#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
parser_laptop.py - Парсер ноутбуков v2.0
Специализированный парсер для поиска ноутбуков в госзакупках

КЛЮЧЕВЫЕ ВОЗМОЖНОСТИ v2.0:
- Трёхэтапное получение данных: список закупок -> детали лотов -> позиции и параметры
- Полная схема БД с таблицами lots, positions, params
- Целенаправленный поиск ноутбуков по КТРУ и ключевым словам
- Фильтрация по размеру диагонали экрана
- Многопоточность с chunked периодами
- Теплое сохранение данных
- Экспорт в CSV и Excel (3 листа/файла)
"""

import os
import sys
import json
import time
import csv
import sqlite3
import logging
import argparse
import signal
import threading
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple, Any
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from dataclasses import dataclass

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Опциональные зависимости
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    XLSX_AVAILABLE = True
except ImportError:
    XLSX_AVAILABLE = False

try:
    import requests_cache
    CACHE_AVAILABLE = True
except ImportError:
    CACHE_AVAILABLE = False

try:
    from tqdm import tqdm
    TQDM_AVAILABLE = True
except ImportError:
    TQDM_AVAILABLE = False

# БАЗОВЫЕ КОНСТАНТЫ API
BASE_URL = "https://v2.gosplan.info/fz44"
LIST_ENDPOINT = "/purchases"
HEADERS = {"Accept": "application/json", "Accept-Encoding": "gzip"}

# КОНФИГУРАЦИЯ НОУТБУКОВ (захардкожена)
LAPTOP_CONFIG = {
    'name': 'Ноутбук',
    'ktru': [
        '26.20.11.110-00000138',
        '26.20.11.110-00000139', 
        '26.20.11.110-00000140',
        '26.20.11.110-00000141',
        '26.20.11.110-00000142',
        '26.20.11.110-00000143',
        '26.20.11.110-00000144',
        '26.20.11.110-00000145',
        '26.20.11.110-00000165'
    ],
    'include_keywords': [
        'ноутбук', 'ультрабук', 'laptop', 
        'портативная пэвм', 'портативный персональный компьютер',
        'мобильное персональное рабочее место', 'мобильное арм'
    ],
    'exclude_keywords': [
        'планшет', 'системный блок', 'моноблок', 'aio'
    ],
    'required_param': 'размер диагонали экрана'
}

# ЛИМИТЫ И НАСТРОЙКИ API
API_LIMITS = {
    "v2.gosplan.info": 600,
    "test.gosplan.info": 10,
}

DEFAULT_PAGE_LIMIT = 100
MAX_PAGINATION_OFFSET = 10_000
MAX_THREADS = min(2 * os.cpu_count(), 20)
EXPORT_FILENAME_MASK = "%Y%m%d"

# Глобальные переменные
interrupted = False
logger = None

def setup_logging(debug: bool = False):
    """Настройка логирования"""
    global logger
    
    log_filename = datetime.now().strftime("laptop_parser_%Y%m%d_%H%M.log")
    level = logging.DEBUG if debug else logging.INFO
    
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    logger = logging.getLogger(__name__)
    logger.info(f"Парсер ноутбуков v2.0 запущен. Лог: {log_filename}")

def signal_handler(signum, frame):
    """Обработчик прерывания"""
    global interrupted
    interrupted = True
    print("\nПолучен сигнал прерывания...")
    if logger:
        logger.info("Получен сигнал прерывания")

signal.signal(signal.SIGINT, signal_handler)
signal.signal(signal.SIGTERM, signal_handler)

@dataclass
class RequestStats:
    """Статистика запросов"""
    requests_count: int = 0
    pages_fetched: int = 0
    lots_total: int = 0
    lots_found: int = 0
    positions_total: int = 0
    params_total: int = 0
    errors_count: int = 0
    start_time: float = 0

class AdaptiveRateLimiter:
    """Адаптивный rate limiter"""
    
    def __init__(self, base_url: str):
        self.latencies = []
        self.current_delay = 0.5
        self.lock = threading.Lock()
        
        for url_pattern, limit in API_LIMITS.items():
            if url_pattern in base_url:
                self.requests_per_minute = limit
                break
        else:
            self.requests_per_minute = 600
        
        self.window_start = time.time()
        self.requests_in_window = 0
        
        logger.info(f"Rate limiter: {self.requests_per_minute} запросов/мин")
    
    def record_request(self, latency: float):
        """Записать запрос и латентность"""
        with self.lock:
            current_time = time.time()
            
            if current_time - self.window_start >= 60:
                self.window_start = current_time
                self.requests_in_window = 0
            
            self.requests_in_window += 1
            self.latencies.append(latency)
            
            if len(self.latencies) > 100:
                self.latencies.pop(0)
    
    def get_delay(self) -> float:
        """Получить задержку"""
        with self.lock:
            if self.requests_in_window >= self.requests_per_minute:
                return 60.0
            
            if len(self.latencies) >= 10:
                sorted_latencies = sorted(self.latencies)
                p99_idx = int(len(sorted_latencies) * 0.99)
                p99 = sorted_latencies[p99_idx]
                
                if p99 > 0.5:
                    self.current_delay = min(self.current_delay + 0.1, 2.0)
                elif p99 < 0.2:
                    self.current_delay = max(self.current_delay - 0.1, 0.1)
            
            return self.current_delay

class HTTPClient:
    """HTTP клиент с кэшированием"""
    
    def __init__(self, cache_ttl_hours: int = 24):
        self.base_url = BASE_URL
        self.session = self._create_session()
        self.rate_limiter = AdaptiveRateLimiter(self.base_url)
        
        if CACHE_AVAILABLE:
            requests_cache.install_cache(
                cache_name='laptop_cache',
                backend='sqlite',
                expire_after=cache_ttl_hours * 3600
            )
            logger.info(f"HTTP кэш включен: TTL {cache_ttl_hours}ч")
        else:
            logger.warning("requests-cache недоступен")
    
    def _create_session(self):
        """Создание HTTP сессии"""
        session = requests.Session()
        
        retry = Retry(
            total=5,
            backoff_factor=1,
            status_forcelist=[500, 502, 503, 504],
            allowed_methods=["GET"]
        )
        
        adapter = HTTPAdapter(max_retries=retry)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        session.headers.update(HEADERS)
        
        return session
    
    def get(self, endpoint: str, params: dict = None) -> Optional[dict]:
        """GET запрос с rate limiting"""
        global interrupted
        
        if interrupted:
            return None
        
        url = f"{self.base_url}{endpoint}"
        max_retries = 5
        
        for attempt in range(max_retries):
            start_time = time.monotonic()
            
            try:
                delay = self.rate_limiter.get_delay()
                if delay > 0:
                    time.sleep(delay)
                
                resp = self.session.get(url, params=params, timeout=60)
                latency = time.monotonic() - start_time
                self.rate_limiter.record_request(latency)
                
                if resp.status_code == 200:
                    return resp.json()
                elif resp.status_code == 404:
                    return None
                elif resp.status_code == 429:
                    delay = min(60 * (2 ** attempt), 300)
                    logger.warning(f"HTTP 429, пауза {delay}s")
                    time.sleep(delay)
                    continue
                elif 500 <= resp.status_code < 600:
                    delay = min(2 ** attempt, 16)
                    logger.warning(f"HTTP {resp.status_code}, back-off {delay}s")
                    time.sleep(delay)
                    continue
                else:
                    logger.error(f"HTTP {resp.status_code} для {url}")
                    return None
                    
            except requests.RequestException as e:
                if attempt == max_retries - 1:
                    logger.error(f"Окончательная ошибка для {url}: {e}")
                    return None
                
                delay = min(2 ** attempt, 16)
                logger.warning(f"Сетевая ошибка, повтор через {delay}s: {e}")
                time.sleep(delay)
                continue
        
        return None

class Database:
    """Управление базой данных для ноутбуков v2.0"""
    
    def __init__(self, db_path: str = "laptop_parser.db"):
        self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        """Инициализация схемы БД v2.0"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Таблица лотов
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS lots (
                purchase_id TEXT NOT NULL,
                lot_number TEXT NOT NULL,
                lot_name TEXT,
                okpd2 TEXT,
                ktru TEXT,
                qty REAL,
                uom TEXT,
                start_price_rub REAL,
                date_bid_start TEXT,
                date_bid_end TEXT,
                delivery_deadline TEXT,
                customer_name TEXT,
                customer_inn TEXT,
                published_at TEXT,
                contract_price REAL,
                winner_name TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (purchase_id, lot_number)
            )
        ''')
        
        # Таблица позиций
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS positions (
                purchase_id TEXT NOT NULL,
                lot_number TEXT NOT NULL,
                position_number TEXT NOT NULL,
                item_name TEXT,
                qty REAL,
                uom TEXT,
                unit_price_rub REAL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (purchase_id, lot_number, position_number),
                FOREIGN KEY (purchase_id, lot_number) REFERENCES lots(purchase_id, lot_number)
            )
        ''')
        
        # Обновленная таблица параметров с position_number
        cursor.execute('DROP TABLE IF EXISTS params')
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS params (
                purchase_id TEXT NOT NULL,
                lot_number TEXT NOT NULL,
                position_number TEXT,
                param_name TEXT NOT NULL,
                param_value TEXT,
                param_unit TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (purchase_id, lot_number) REFERENCES lots(purchase_id, lot_number)
            )
        ''')
        
        # KV store для last_seen
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS kv_store (
                key TEXT PRIMARY KEY,
                value TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Индексы
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_lots_purchase ON lots(purchase_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_lots_price ON lots(start_price_rub)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_lots_date ON lots(published_at)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_positions_lot ON positions(purchase_id, lot_number)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_positions_price ON positions(unit_price_rub)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_params_pos ON params(purchase_id, lot_number, position_number)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_params_name ON params(param_name)')
        
        # FTS для поиска
        try:
            cursor.execute('''
                CREATE VIRTUAL TABLE IF NOT EXISTS lots_fts USING fts5(
                    purchase_id, lot_name, customer_name,
                    tokenize='unicode61 remove_diacritics 1'
                )
            ''')
            logger.info("FTS5 индекс создан")
        except Exception as e:
            logger.warning(f"FTS5 недоступен: {e}")
        
        conn.commit()
        conn.close()
        
        logger.info(f"База данных v2.0 инициализирована: {self.db_path}")
    
    def get_last_seen(self) -> Optional[str]:
        """Получить last_seen дату"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT value FROM kv_store WHERE key = ?', ('last_seen',))
        result = cursor.fetchone()
        conn.close()
        return result[0] if result else None
    
    def set_last_seen(self, date_str: str):
        """Установить last_seen дату"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO kv_store (key, value)
            VALUES (?, ?)
        ''', ('last_seen', date_str))
        conn.commit()
        conn.close()
        logger.debug(f"Обновлен last_seen: {date_str}")
    
    def save_lots(self, lots_data: List[Dict]) -> int:
        """Сохранение лотов ноутбуков"""
        if not lots_data:
            return 0
        
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        saved_count = 0
        
        for lot_data in lots_data:
            try:
                # Сохраняем основные данные лота
                cursor.execute('''
                    INSERT OR REPLACE INTO lots 
                    (purchase_id, lot_number, lot_name, okpd2, ktru, qty, uom,
                     start_price_rub, date_bid_start, date_bid_end, delivery_deadline,
                     customer_name, customer_inn, published_at, contract_price, winner_name)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    lot_data.get('purchase_id', ''),
                    lot_data.get('lot_number', ''),
                    lot_data.get('lot_name', ''),
                    lot_data.get('okpd2', ''),
                    lot_data.get('ktru', ''),
                    lot_data.get('qty'),
                    lot_data.get('uom', ''),
                    lot_data.get('start_price_rub'),
                    lot_data.get('date_bid_start', ''),
                    lot_data.get('date_bid_end', ''),
                    lot_data.get('delivery_deadline', ''),
                    lot_data.get('customer_name', ''),
                    lot_data.get('customer_inn', ''),
                    lot_data.get('published_at', ''),
                    lot_data.get('contract_price'),
                    lot_data.get('winner_name', '')
                ))
                
                # Обновляем FTS индекс если доступен
                try:
                    cursor.execute('''
                        INSERT OR REPLACE INTO lots_fts (purchase_id, lot_name, customer_name)
                        VALUES (?, ?, ?)
                    ''', (
                        lot_data.get('purchase_id', ''),
                        lot_data.get('lot_name', ''),
                        lot_data.get('customer_name', '')
                    ))
                except:
                    pass  # FTS недоступен
                
                saved_count += 1
                
            except Exception as e:
                logger.error(f"Ошибка сохранения лота {lot_data.get('purchase_id', '')}: {e}")
        
        conn.commit()
        conn.close()
        
        if saved_count > 0:
            logger.info(f"Сохранено {saved_count} лотов ноутбуков")
        
        return saved_count
    
    def save_positions(self, position_rows: List[Dict]) -> int:
        """Сохранение позиций"""
        if not position_rows:
            return 0
            
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        count = 0
        
        for row in position_rows:
            try:
                cursor.execute('''
                    INSERT OR REPLACE INTO positions
                    (purchase_id, lot_number, position_number, item_name,
                     qty, uom, unit_price_rub)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    row['purchase_id'], row['lot_number'], row['position_number'],
                    row['item_name'], row['qty'], row['uom'], row['unit_price_rub']
                ))
                count += 1
            except Exception as e:
                logger.error(f"Ошибка сохранения позиции {row['purchase_id']}[{row['lot_number']}]: {e}")
        
        conn.commit()
        conn.close()
        
        if count > 0:
            logger.info(f"Сохранено {count} позиций ноутбуков")
        return count
    
    def save_params(self, params_rows: List[Dict]) -> int:
        """Сохранение параметров"""
        if not params_rows:
            return 0
            
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        count = 0
        
        for row in params_rows:
            try:
                cursor.execute('''
                    INSERT INTO params
                    (purchase_id, lot_number, position_number, param_name,
                     param_value, param_unit)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (
                    row['purchase_id'], row['lot_number'], row.get('position_number'),
                    row['param_name'], row['param_value'], row.get('param_unit', '')
                ))
                count += 1
            except Exception as e:
                logger.error(f"Ошибка сохранения параметра {row['purchase_id']}[{row['lot_number']}][{row.get('position_number')}]: {e}")
        
        conn.commit()
        conn.close()
        
        if count > 0:
            logger.info(f"Сохранено {count} параметров ноутбуков")
        return count

def normalize_text(text: str) -> str:
    """Нормализация текста"""
    if not text:
        return ""
    return " ".join(text.lower().strip().split())

def is_laptop(lot: Dict) -> Tuple[bool, List[str]]:
    """
    Проверка является ли лот ноутбуком
    
    Проверяет:
    1. Include keywords (должно быть хотя бы одно)
    2. Exclude keywords (не должно быть ни одного) 
    3. Обязательный параметр "размер диагонали экрана"
    
    Возвращает: (is_laptop, reasons)
    """
    reasons = []
    
    # Собираем весь текст для поиска
    searchable_text = ""
    if lot.get('lot_name') or lot.get('lotName'):
        searchable_text += " " + str(lot.get('lot_name') or lot.get('lotName'))
    if lot.get('object_info'):
        searchable_text += " " + str(lot['object_info'])
    
    # Добавляем текст из характеристик
    def _extract_text_from_nested(data):
        nonlocal searchable_text
        if isinstance(data, dict):
            for key, value in data.items():
                if isinstance(value, str):
                    searchable_text += " " + value
                elif isinstance(value, (dict, list)):
                    _extract_text_from_nested(value)
        elif isinstance(data, list):
            for item in data:
                _extract_text_from_nested(item)
        elif isinstance(data, str):
            searchable_text += " " + data
    
    for section in ['characteristics', 'specifications', 'parameters', 'items']:
        if section in lot:
            _extract_text_from_nested(lot[section])
    
    searchable_text = normalize_text(searchable_text)
    
    # 1. Проверка include keywords (должно быть хотя бы одно)
    found_include = False
    for keyword in LAPTOP_CONFIG['include_keywords']:
        if normalize_text(keyword) in searchable_text:
            found_include = True
            reasons.append(f"include:{keyword}")
            break
    
    if not found_include:
        return False, ["no_include_keywords"]
    
    # 2. Проверка exclude keywords (не должно быть ни одного)
    for keyword in LAPTOP_CONFIG['exclude_keywords']:
        if normalize_text(keyword) in searchable_text:
            return False, [f"exclude:{keyword}"]
    
    # 3. Проверка обязательного параметра "размер диагонали экрана"
    found_screen_size = False
    
    def _find_screen_size_param(data):
        nonlocal found_screen_size
        if isinstance(data, dict):
            # Ищем пары parameter/value
            if 'parameter' in data or 'param_name' in data or 'name' in data:
                param_name = data.get('parameter') or data.get('param_name') or data.get('name', '')
                param_value = data.get('value') or data.get('param_value') or data.get('val', '')
                
                if param_name and param_value:
                    normalized_param = normalize_text(str(param_name))
                    
                    # Ищем параметр диагонали экрана
                    if LAPTOP_CONFIG['required_param'] in normalized_param or 'диагональ экрана' in normalized_param:
                        if param_value and str(param_value).strip():
                            found_screen_size = True
                            reasons.append(f"screen_size:{param_value}")
            
            # Рекурсивно ищем дальше
            for key, value in data.items():
                _find_screen_size_param(value)
        
        elif isinstance(data, (list, tuple)):
            for item in data:
                _find_screen_size_param(item)
    
    # Ищем параметр диагонали во всем лоте
    _find_screen_size_param(lot)
    
    if not found_screen_size:
        return False, ["missing_screen_size_param"]
    
    return True, reasons

def batched(iterable, n):
    """Разбивка списка на батчи"""
    import itertools
    iterator = iter(iterable)
    while True:
        batch = list(itertools.islice(iterator, n))
        if not batch:
            break
        yield batch

class LaptopParser:
    """Парсер ноутбуков v2.0 с трёхэтапным получением данных"""
    
    def __init__(self, threads: int = MAX_THREADS, chunk_hours: int = 24, 
                 show_progress: bool = False, min_price: float = None, max_price: float = None,
                 regions: List[int] = None):
        self.http_client = HTTPClient()
        self.database = Database()
        self.stats = RequestStats()
        self.threads = max(1, min(threads, MAX_THREADS))
        self.chunk_hours = max(1, min(chunk_hours, 24))
        self.show_progress = show_progress and TQDM_AVAILABLE
        self.min_price = min_price
        self.max_price = max_price
        self.regions = regions or []
        
        logger.info(f"Парсер ноутбуков v2.0 инициализирован (потоков: {self.threads}, chunk: {self.chunk_hours}ч)")
    
    def fetch_purchase_details(self, purchase_number: str) -> Optional[dict]:
        """
        Получить детальную информацию о закупке по её номеру.
        Возвращает словарь с ключами: purchase, lots, positions
        """
        if interrupted:
            return None
            
        result = {
            'purchase': None,
            'lots': [],
            'positions': []
        }
        
        # 1. Получаем базовую информацию о закупке
        self.stats.requests_count += 1
        purchase_resp = self.http_client.get(f"/purchase/{purchase_number}")
        if purchase_resp:
            result['purchase'] = purchase_resp
        
        # 2. Получаем лоты
        self.stats.requests_count += 1
        lots_resp = self.http_client.get(f"/purchase/{purchase_number}/lots")
        if lots_resp:
            if isinstance(lots_resp, list):
                result['lots'] = lots_resp
            elif isinstance(lots_resp, dict) and 'lots' in lots_resp:
                result['lots'] = lots_resp['lots']
        
        # 3. Получаем позиции
        self.stats.requests_count += 1
        pos_resp = self.http_client.get(f"/purchase/{purchase_number}/positions")
        if pos_resp:
            if isinstance(pos_resp, list):
                result['positions'] = pos_resp
            elif isinstance(pos_resp, dict) and 'positions' in pos_resp:
                result['positions'] = pos_resp['positions']
        else:
            # Пробуем альтернативный эндпоинт
            self.stats.requests_count += 1
            items_resp = self.http_client.get(f"/purchase/{purchase_number}/items")
            if items_resp:
                if isinstance(items_resp, list):
                    result['positions'] = items_resp
                elif isinstance(items_resp, dict) and 'items' in items_resp:
                    result['positions'] = items_resp['items']
        
        # Проверяем что хоть что-то получили
        if not result['purchase'] and not result['lots'] and not result['positions']:
            return None
            
        return result
    
    def _create_pseudo_lot(self, purchase: Dict) -> List[Dict]:
        """Создание псевдо-лота из метаданных закупки (fallback)"""
        return [{
            'lotNumber': '1',
            'lotName': purchase.get('object_info', ''),
            'okpd2': purchase.get('okpd2', [''])[0] if purchase.get('okpd2') else '',
            'ktru': purchase.get('ktru', [''])[0] if purchase.get('ktru') else '',
            'price': purchase.get('max_price') or purchase.get('start_price'),
            'quantity': 1,
            'unit': {'shortName': 'шт'},
            'characteristics': []
        }]
    
    def _create_time_chunks(self, start_date: datetime, end_date: datetime) -> List[Tuple[datetime, datetime]]:
        """Создание временных чанков"""
        chunks = []
        step = timedelta(hours=self.chunk_hours)
        
        curr_start = start_date
        while curr_start < end_date:
            curr_end = min(curr_start + step - timedelta(seconds=1), end_date)
            chunks.append((curr_start, curr_end))
            curr_start = curr_end + timedelta(seconds=1)
        
        logger.info(f"Создано {len(chunks)} временных чанков по {self.chunk_hours}ч")
        return chunks
    
    def _extract_lot_metadata(self, lot: Dict, purchase_data: Dict) -> Dict:
        """Извлечение метаданных лота (без параметров)"""
        # Информация о заказчике
        customers = purchase_data.get('customers', [])
        if customers and isinstance(customers[0], dict):
            customer_name = customers[0].get('name', '')
            customer_inn = customers[0].get('inn', '')
        elif customers:
            customer_name = str(customers[0])
            customer_inn = ''
        else:
            customer_name = ''
            customer_inn = ''
        
        # Извлекаем цену
        start_price = self._extract_number(purchase_data, ['max_price', 'startPrice']) or self._extract_number(lot, ['price'])
        
        return {
            'purchase_id': purchase_data.get('purchase_number', ''),
            'lot_number': str(lot.get('lotNumber', lot.get('lot_number', '1'))),
            'lot_name': lot.get('lotName', lot.get('name', '')),
            'okpd2': lot.get('okpd2', lot.get('okpd2Code', '')),
            'ktru': lot.get('ktru', lot.get('ktruCode', lot.get('code', ''))),
            'qty': self._extract_number(lot, ['quantity', 'qty', 'value']),
            'uom': lot.get('unitName', lot.get('unit', {}).get('shortName', '') if isinstance(lot.get('unit'), dict) else ''),
            'start_price_rub': start_price,
            'date_bid_start': purchase_data.get('startOfReception', purchase_data.get('collecting_started_at', '')),
            'date_bid_end': purchase_data.get('collecting_finished_at', purchase_data.get('endOfReception', '')),
            'delivery_deadline': lot.get('delivery_end_date', lot.get('deliveryDate', '')),
            'customer_name': customer_name,
            'customer_inn': customer_inn,
            'published_at': purchase_data.get('published_at', ''),
            'contract_price': self._extract_number(purchase_data, ['contract_price', 'contractPrice']),
            'winner_name': lot.get('winner', {}).get('name', '') if isinstance(lot.get('winner'), dict) else ''
        }
    
    def _extract_lot_dynamic_params(self, lot: Dict) -> List[Dict]:
        """Извлечение динамических параметров лота"""
        params = []
        
        def _extract_params_recursive(data, prefix=""):
            if isinstance(data, dict):
                if 'parameter' in data and 'value' in data:
                    param_name = str(data['parameter'])
                    param_value = str(data['value'])
                    param_unit = str(data.get('unit', ''))
                    if param_name and param_value:
                        full_name = f"{prefix}{param_name}" if prefix else param_name
                        params.append({
                            'param_name': full_name,
                            'param_value': param_value,
                            'param_unit': param_unit
                        })
                
                for key, value in data.items():
                    if key not in ['parameter', 'value', 'unit']:
                        new_prefix = f"{prefix}{key}." if prefix else f"{key}."
                        _extract_params_recursive(value, new_prefix)
            
            elif isinstance(data, (list, tuple)):
                for i, item in enumerate(data):
                    _extract_params_recursive(item, f"{prefix}[{i}].")
        
        for section in ['characteristics', 'specifications', 'parameters', 'items']:
            if section in lot:
                _extract_params_recursive(lot[section], f"lot.{section}.")
        
        return params
    
    def _extract_positions_data(self, positions: List[Dict], purchase_id: str, lot_number: str) -> Tuple[List[Dict], List[Dict]]:
        """Извлечение данных позиций и их параметров"""
        position_rows = []
        params_rows = []
        
        for pos_idx, position in enumerate(positions):
            # Извлекаем данные позиции
            position_number = str(position.get('position_number', position.get('positionNumber', pos_idx + 1)))
            
            position_rows.append({
                'purchase_id': purchase_id,
                'lot_number': lot_number,
                'position_number': position_number,
                'item_name': position.get('name', position.get('item_name', '')),
                'qty': self._extract_number(position, ['quantity', 'qty']),
                'uom': position.get('unit', position.get('unitName', '')),
                'unit_price_rub': self._extract_number(position, ['initial_price', 'price', 'unitPrice'])
            })
            
            self.stats.positions_total += 1
            
            # Извлекаем параметры позиции
            for param in position.get('parameters', []):
                param_name = param.get('name', param.get('parameter', ''))
                param_value = param.get('value', '')
                param_unit = param.get('unit', '')
                
                if param_name and param_value:
                    params_rows.append({
                        'purchase_id': purchase_id,
                        'lot_number': lot_number,
                        'position_number': position_number,
                        'param_name': param_name,
                        'param_value': str(param_value),
                        'param_unit': param_unit
                    })
                    self.stats.params_total += 1
        
        return position_rows, params_rows
    
    def _extract_number(self, data: Dict, possible_keys: List[str]) -> Optional[float]:
        """Извлечение числового значения"""
        for key in possible_keys:
            if key in data and data[key] is not None:
                try:
                    return float(data[key])
                except (ValueError, TypeError):
                    continue
        return None
    
    def _fetch_laptops(self, start_date: datetime, end_date: datetime) -> int:
        """Получение ноутбуков с chunking и трёхэтапным получением"""
        
        # Создаем временные чанки
        time_chunks = self._create_time_chunks(start_date, end_date)
        
        # Разбиваем коды КТРУ на батчи по 10
        ktru_codes = LAPTOP_CONFIG['ktru']
        code_batches = list(batched(ktru_codes, 10))
        
        total_combinations = len(time_chunks) * len(code_batches)
        
        logger.info(f"Обработка {len(time_chunks)} временных чанков × {len(code_batches)} батчей кодов = {total_combinations} комбинаций")
        
        # Создаем итератор для progress bar
        if self.show_progress:
            combinations_iter = tqdm(
                [(chunk, batch) for chunk in time_chunks for batch in code_batches],
                desc="Поиск ноутбуков",
                unit="batch"
            )
        else:
            combinations_iter = [(chunk, batch) for chunk in time_chunks for batch in code_batches]
        
        def process_combination(chunk_and_batch):
            """Обработка одной комбинации с трёхэтапным получением"""
            (chunk_start, chunk_end), code_batch = chunk_and_batch
            
            if interrupted:
                return 0
            
            try:
                purchases = self._fetch_period_recursive(code_batch, chunk_start, chunk_end)
                
                if not purchases:
                    return 0
                
                # Обрабатываем каждую закупку
                for purchase in purchases:
                    if interrupted:
                        break
                    
                    purchase_number = purchase.get('purchase_number')
                    if not purchase_number:
                        continue
                    
                    # Получаем детальную информацию
                    details = self.fetch_purchase_details(purchase_number)
                    if not details:
                        # Fallback на псевдо-лот
                        lots_to_process = self._create_pseudo_lot(purchase)
                    else:
                        # Объединяем данные
                        if details['purchase']:
                            purchase.update(details['purchase'])
                        lots_to_process = details['lots'] if details['lots'] else self._create_pseudo_lot(purchase)
                    
                    # Буферы для batch сохранения
                    extracted_lot_rows = []
                    position_rows = []
                    params_rows = []
                    
                    # Обрабатываем каждый лот
                    for lot_idx, lot in enumerate(lots_to_process):
                        self.stats.lots_total += 1
                        
                        # Проверяем является ли лот ноутбуком
                        laptop_match, match_reasons = is_laptop(lot)
                        
                        if not laptop_match:
                            logger.debug(f"Лот {purchase_number}[{lot_idx}] не ноутбук: {match_reasons}")
                            continue
                        
                        # Извлекаем метаданные лота
                        lot_meta = self._extract_lot_metadata(lot, purchase)
                        
                        # Проверяем ценовые ограничения
                        if self.min_price is not None and lot_meta['start_price_rub'] is not None and lot_meta['start_price_rub'] < self.min_price:
                            logger.debug(f"Лот {purchase_number}[{lot_idx}] ниже min_price")
                            continue
                            
                        if self.max_price is not None and lot_meta['start_price_rub'] is not None and lot_meta['start_price_rub'] > self.max_price:
                            logger.debug(f"Лот {purchase_number}[{lot_idx}] выше max_price")
                            continue
                        
                        self.stats.lots_found += 1
                        extracted_lot_rows.append(lot_meta)
                        
                        # Извлекаем параметры лота
                        lot_params = self._extract_lot_dynamic_params(lot)
                        for param in lot_params:
                            params_rows.append({
                                'purchase_id': lot_meta['purchase_id'],
                                'lot_number': lot_meta['lot_number'],
                                'position_number': None,
                                **param
                            })
                        
                        # Извлекаем позиции если есть
                        if details and details['positions']:
                            pos_rows, pos_params = self._extract_positions_data(
                                details['positions'], 
                                lot_meta['purchase_id'], 
                                lot_meta['lot_number']
                            )
                            position_rows.extend(pos_rows)
                            params_rows.extend(pos_params)
                        
                        logger.debug(f"НАЙДЕН лот ноутбука {purchase_number}[{lot_idx}]: {match_reasons}")
                    
                    # Теплое сохранение в БД
                    if extracted_lot_rows:
                        self.database.save_lots(extracted_lot_rows)
                        if position_rows:
                            self.database.save_positions(position_rows)
                        if params_rows:
                            self.database.save_params(params_rows)
                        
                        logger.debug(f"Загружена закупка {purchase_number}: {len(extracted_lot_rows)} лотов, {len(position_rows)} позиций, {len(params_rows)} параметров")
                
                return len(purchases)
                
            except Exception as e:
                logger.error(f"Ошибка обработки комбинации: {e}")
                self.stats.errors_count += 1
                
                # Авто-retry при серверных ошибках
                if "5" in str(e) or "429" in str(e):
                    logger.warning("Ошибки сервера, пауза 120 сек...")
                    time.sleep(120)
                    try:
                        purchases = self._fetch_period_recursive(code_batch, chunk_start, chunk_end)
                        return len(purchases) if purchases else 0
                    except Exception as retry_e:
                        logger.error(f"Повторная ошибка: {retry_e}")
                
                return 0
        
        # Обрабатываем комбинации
        total_purchases = 0
        
        if self.threads == 1:
            for combination in combinations_iter:
                if interrupted:
                    break
                count = process_combination(combination)
                total_purchases += count
        else:
            with ThreadPoolExecutor(max_workers=self.threads) as executor:
                future_to_combination = {
                    executor.submit(process_combination, combination): combination 
                    for combination in combinations_iter
                }
                
                for future in as_completed(future_to_combination):
                    if interrupted:
                        break
                    
                    try:
                        count = future.result()
                        total_purchases += count
                    except Exception as e:
                        logger.error(f"Ошибка в потоке: {e}")
                        self.stats.errors_count += 1
        
        logger.info(f"Загрузка ноутбуков завершена: {total_purchases} закупок")
        return total_purchases
    
    def _fetch_period_recursive(self, code_batch: List[str], 
                              start_date: datetime, end_date: datetime) -> List[Dict]:
        """Рекурсивная загрузка периода (только список закупок)"""
        if interrupted:
            return []
        
        start_iso = start_date.strftime("%Y-%m-%dT%H:%M:%SZ")
        end_iso = end_date.strftime("%Y-%m-%dT%H:%M:%SZ")
        
        all_data = []
        offset = 0
        
        while True:
            if interrupted:
                break
            
            # Формируем параметры запроса
            params = {
                "limit": DEFAULT_PAGE_LIMIT,
                "offset": offset,
                "published_after": start_iso,
                "published_before": end_iso,
            }
            
            # Добавляем коды КТРУ
            for i, code in enumerate(code_batch[:10]):
                params[f"classifiers[{i}]"] = code
            
            # Добавляем регионы если указаны
            if self.regions:
                params["region"] = self.regions
            
            self.stats.requests_count += 1
            page_data = self.http_client.get(LIST_ENDPOINT, params)
            
            if not page_data:
                break
            
            self.stats.pages_fetched += 1
            
            # Добавляем в результат (детали будут получены отдельно)
            if isinstance(page_data, list):
                all_data.extend(page_data)
            elif isinstance(page_data, dict) and 'purchases' in page_data:
                all_data.extend(page_data['purchases'])
            
            if len(page_data) < DEFAULT_PAGE_LIMIT:
                break
            
            offset += DEFAULT_PAGE_LIMIT
            
            # Если достигли предела пагинации - разбиваем период
            if offset >= MAX_PAGINATION_OFFSET:
                logger.info(f"Достигнут предел пагинации, разбиваем период")
                
                total_seconds = (end_date - start_date).total_seconds()
                mid_date = start_date + timedelta(seconds=total_seconds / 2)
                
                first_half = self._fetch_period_recursive(code_batch, start_date, mid_date - timedelta(seconds=1))
                second_half = self._fetch_period_recursive(code_batch, mid_date, end_date)
                
                all_data.extend(first_half)
                all_data.extend(second_half)
                break
        
        return all_data
    
    def parse_laptops(self, start_date: datetime, end_date: datetime) -> int:
        """Главный метод парсинга ноутбуков v2.0"""
        self.stats.start_time = time.time()
        
        logger.info(f"Начинаем поиск ноутбуков с {start_date} по {end_date}")
        print(f"Поиск ноутбуков с {start_date.date()} по {end_date.date()}")
        
        # Поиск и сохранение (трёхэтапное получение)
        total_purchases = self._fetch_laptops(start_date, end_date)
        
        # Подсчитываем сохраненные лоты из БД
        conn = sqlite3.connect(self.database.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(*) FROM lots')
        saved_count = cursor.fetchone()[0]
        conn.close()
        
        if saved_count > 0:
            print(f"НАЙДЕНО ноутбуков: {saved_count} (обработано {total_purchases} закупок)")
        else:
            print(f"Ноутбуки не найдены (обработано {total_purchases} закупок)")
        
        # Обновляем last_seen
        self.database.set_last_seen(end_date.strftime("%Y-%m-%dT%H:%M:%SZ"))
        
        return saved_count

class LaptopExporter:
    """Экспорт данных ноутбуков v2.0"""
    
    def __init__(self, database: Database):
        self.database = database
    
    def get_lots_export_data(self) -> Tuple[List[Dict], List[str]]:
        """Получение данных лотов для экспорта"""
        conn = sqlite3.connect(self.database.db_path)
        cursor = conn.cursor()
        
        # Получаем данные лотов
        cursor.execute('''
            SELECT * FROM lots
            ORDER BY created_at DESC
        ''')
        
        columns = [desc[0] for desc in cursor.description]
        results = cursor.fetchall()
        
        conn.close()
        
        # Формируем список словарей
        export_data = []
        for row in results:
            item = dict(zip(columns, row))
            export_data.append(item)
        
        return export_data, columns
    
    def get_positions_export_data(self) -> Tuple[List[Dict], List[str]]:
        """Получение данных позиций для экспорта"""
        conn = sqlite3.connect(self.database.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM positions
            ORDER BY purchase_id, lot_number, position_number
        ''')
        
        columns = [desc[0] for desc in cursor.description]
        results = cursor.fetchall()
        
        conn.close()
        
        export_data = []
        for row in results:
            item = dict(zip(columns, row))
            export_data.append(item)
        
        return export_data, columns
    
    def get_params_export_data(self) -> Tuple[List[Dict], List[str]]:
        """Получение данных параметров для экспорта"""
        conn = sqlite3.connect(self.database.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM params
            ORDER BY purchase_id, lot_number, position_number, param_name
        ''')
        
        columns = [desc[0] for desc in cursor.description]
        results = cursor.fetchall()
        
        conn.close()
        
        export_data = []
        for row in results:
            item = dict(zip(columns, row))
            export_data.append(item)
        
        return export_data, columns
    
    def export_to_csv(self, base_filename: str = None, export_positions: bool = True, export_params: bool = True) -> List[str]:
        """Экспорт в CSV файлы"""
        if not base_filename:
            base_filename = f"laptop_{datetime.now().strftime(EXPORT_FILENAME_MASK)}"
        
        exported_files = []
        
        # 1. Экспорт лотов
        lots_data, lots_columns = self.get_lots_export_data()
        lots_filename = f"{base_filename}_lots.csv"
        
        if lots_data:
            with open(lots_filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=lots_columns)
                writer.writeheader()
                writer.writerows(lots_data)
            
            print(f"Экспорт лотов CSV: {lots_filename} ({len(lots_data)} записей)")
            exported_files.append(lots_filename)
        else:
            print("ВНИМАНИЕ: Нет данных лотов для экспорта")
        
        # 2. Экспорт позиций
        if export_positions:
            positions_data, positions_columns = self.get_positions_export_data()
            if positions_data:
                positions_filename = f"{base_filename}_positions.csv"
                with open(positions_filename, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=positions_columns)
                    writer.writeheader()
                    writer.writerows(positions_data)
                
                print(f"Экспорт позиций CSV: {positions_filename} ({len(positions_data)} записей)")
                exported_files.append(positions_filename)
        
        # 3. Экспорт параметров
        if export_params:
            params_data, params_columns = self.get_params_export_data()
            if params_data:
                params_filename = f"{base_filename}_params.csv"
                with open(params_filename, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=params_columns)
                    writer.writeheader()
                    writer.writerows(params_data)
                
                print(f"Экспорт параметров CSV: {params_filename} ({len(params_data)} записей)")
                exported_files.append(params_filename)
        
        return exported_files
    
    def export_to_xlsx(self, filename: str = None, export_positions: bool = True, export_params: bool = True) -> str:
        """Экспорт в XLSX с тремя листами"""
        if not XLSX_AVAILABLE:
            raise RuntimeError("openpyxl не установлен. Установите: pip install openpyxl")
        
        if not filename:
            filename = f"laptop_{datetime.now().strftime(EXPORT_FILENAME_MASK)}.xlsx"
        
        wb = Workbook()
        
        # 1. Лист с лотами
        ws_lots = wb.active
        ws_lots.title = "Lots"
        
        lots_data, lots_columns = self.get_lots_export_data()
        
        # Заголовки
        for col, header in enumerate(lots_columns, 1):
            cell = ws_lots.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
        
        # Данные
        for row_idx, item in enumerate(lots_data, 2):
            for col_idx, field in enumerate(lots_columns, 1):
                ws_lots.cell(row=row_idx, column=col_idx, value=item.get(field, ''))
        
        # Автоширина
        self._adjust_column_widths(ws_lots)
        
        # 2. Лист с позициями
        if export_positions:
            ws_positions = wb.create_sheet("Positions")
            positions_data, positions_columns = self.get_positions_export_data()
            
            # Заголовки
            for col, header in enumerate(positions_columns, 1):
                cell = ws_positions.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
            
            # Данные
            for row_idx, item in enumerate(positions_data, 2):
                for col_idx, field in enumerate(positions_columns, 1):
                    ws_positions.cell(row=row_idx, column=col_idx, value=item.get(field, ''))
            
            self._adjust_column_widths(ws_positions)
        
        # 3. Лист с параметрами
        if export_params:
            ws_params = wb.create_sheet("Params")
            params_data, params_columns = self.get_params_export_data()
            
            # Заголовки
            for col, header in enumerate(params_columns, 1):
                cell = ws_params.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
            
            # Данные
            for row_idx, item in enumerate(params_data, 2):
                for col_idx, field in enumerate(params_columns, 1):
                    ws_params.cell(row=row_idx, column=col_idx, value=item.get(field, ''))
            
            self._adjust_column_widths(ws_params)
        
        wb.save(filename)
        
        print(f"Экспорт XLSX: {filename}")
        print(f"  Лотов: {len(lots_data)}")
        if export_positions:
            print(f"  Позиций: {len(positions_data) if 'positions_data' in locals() else 0}")
        if export_params:
            print(f"  Параметров: {len(params_data) if 'params_data' in locals() else 0}")
        
        return filename
    
    def _adjust_column_widths(self, worksheet):
        """Автоподстройка ширины колонок"""
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    cell_length = len(str(cell.value))
                    if cell_length > max_length:
                        max_length = cell_length
                except:
                    pass
            
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

def print_final_statistics(stats: RequestStats):
    """Печать финальной статистики v2.0"""
    runtime_seconds = time.time() - stats.start_time
    runtime_str = f"{int(runtime_seconds // 60)}m {int(runtime_seconds % 60)}s"
    
    print(f"\n" + "="*50)
    print(f"СТАТИСТИКА ПАРСЕРА НОУТБУКОВ v2.0")
    print(f"="*50)
    print(f"HTTP Requests   : {stats.requests_count}")
    print(f"Pages fetched   : {stats.pages_fetched}")
    print(f"Lots total      : {stats.lots_total}")
    print(f"Laptops found   : {stats.lots_found}")
    print(f"Positions total : {stats.positions_total}")
    print(f"Params total    : {stats.params_total}")
    print(f"Errors          : {stats.errors_count}")
    print(f"Runtime         : {runtime_str}")
    print(f"="*50)

def main():
    """Главная функция для парсера ноутбуков v2.0"""
    global interrupted
    
    # CLI аргументы
    parser = argparse.ArgumentParser(description="Парсер ноутбуков v2.0 - расширенный поиск с позициями и параметрами")
    parser.add_argument('--start', type=str, help="Начало периода (YYYY-MM-DD)")
    parser.add_argument('--end', type=str, help="Конец периода (YYYY-MM-DD)")
    parser.add_argument('--min-price', type=float, help="Минимальная цена ноутбука")
    parser.add_argument('--max-price', type=float, help="Максимальная цена ноутбука")
    parser.add_argument('--regions', type=int, nargs='*', help='Коды регионов (77 78 для МСК+СПб)')
    parser.add_argument('--outfile', type=str, help="Базовое имя файла без расширения")
    parser.add_argument('--xlsx', action='store_true', help="Дополнительно экспорт в Excel")
    parser.add_argument('--export-positions', action='store_true', default=True, help="Экспортировать таблицу позиций")
    parser.add_argument('--export-params', action='store_true', default=True, help="Экспортировать таблицу параметров")
    parser.add_argument('--threads', type=int, default=MAX_THREADS, help=f"Число потоков (1-{MAX_THREADS})")
    parser.add_argument('--chunk-hours', type=int, default=24, help='Шаг разрезки периода, часы (1-24)')
    parser.add_argument('--progress', action='store_true', help='Показывать progress bar')
    parser.add_argument('--debug', action='store_true', help="Включить DEBUG логирование")
    
    args = parser.parse_args()
    
    # Валидация chunk-hours
    if args.chunk_hours < 1 or args.chunk_hours > 24:
        print("ОШИБКА: --chunk-hours должен быть от 1 до 24")
        return 1
    
    # Настройка логирования
    setup_logging(args.debug)
    
    print("Компьютер Парсер ноутбуков v2.0")
    print("Расширенный поиск портативных компьютеров")
    print("Трёхэтапное получение: список -> лоты -> позиции+параметры")
    print("Полная схема БД + экспорт в 3 таблицы")
    if not TQDM_AVAILABLE:
        print("ВНИМАНИЕ: Для progress bar установите: pip install tqdm")
    print("")
    
    try:
        # Инициализация
        database = Database()
        parser_instance = LaptopParser(
            threads=args.threads,
            chunk_hours=args.chunk_hours,
            show_progress=args.progress,
            min_price=args.min_price,
            max_price=args.max_price,
            regions=args.regions
        )
        exporter = LaptopExporter(database)
        
        # Отображение конфигурации поиска ноутбуков
        print("Конфигурация поиска ноутбуков:")
        print(f"  КТРУ коды: {len(LAPTOP_CONFIG['ktru'])} кодов")
        print(f"  Include keywords: {', '.join(LAPTOP_CONFIG['include_keywords'][:3])}...")
        print(f"  Exclude keywords: {', '.join(LAPTOP_CONFIG['exclude_keywords'])}")
        print(f"  Обязательный параметр: {LAPTOP_CONFIG['required_param']}")
        if args.min_price:
            print(f"  Минимальная цена: {args.min_price:,.0f} руб")
        if args.max_price:
            print(f"  Максимальная цена: {args.max_price:,.0f} руб")
        if args.regions:
            print(f"  Регионы: {args.regions}")
        print("")
        
        # Определение периода поиска
        if args.start:
            start_date = datetime.strptime(args.start, "%Y-%m-%d")
        else:
            # Используем last_seen или вчера
            last_seen_str = database.get_last_seen()
            if last_seen_str:
                try:
                    start_date = datetime.fromisoformat(last_seen_str.replace('Z', '+00:00'))
                    print(f"Используем last_seen как начальную дату: {start_date.date()}")
                except:
                    start_date = datetime.now() - timedelta(days=1)
                    print(f"Используем вчера: {start_date.date()}")
            else:
                start_date = datetime.now() - timedelta(days=1)
                print(f"last_seen не найден, используем вчера: {start_date.date()}")
        
        if args.end:
            end_date = datetime.strptime(args.end, "%Y-%m-%d")
        else:
            end_date = datetime.now()
        
        # Поиск ноутбуков
        logger.info(f"Запуск поиска ноутбуков с {start_date} по {end_date}")
        result = parser_instance.parse_laptops(start_date, end_date)
        
        # Экспорт
        base_filename = args.outfile if args.outfile else None
        
        # CSV экспорт
        csv_files = exporter.export_to_csv(
            base_filename, 
            export_positions=args.export_positions,
            export_params=args.export_params
        )
        
        # Excel экспорт
        if args.xlsx:
            if XLSX_AVAILABLE:
                xlsx_filename = args.outfile if args.outfile else None
                if xlsx_filename and not xlsx_filename.endswith('.xlsx'):
                    xlsx_filename = xlsx_filename.replace('.csv', '') + '.xlsx'
                exporter.export_to_xlsx(
                    xlsx_filename,
                    export_positions=args.export_positions,
                    export_params=args.export_params
                )
            else:
                print("ОШИБКА: openpyxl не установлен, XLSX экспорт пропущен")
        
        # Финальная статистика
        print_final_statistics(parser_instance.stats)
        
        # Пример SQL запроса
        print("\nПример SQL для объединения таблиц:")
        print("""
SELECT 
    l.purchase_id, 
    l.lot_number, 
    l.lot_name,
    p.position_number, 
    p.item_name, 
    prm.param_name, 
    prm.param_value
FROM lots l
LEFT JOIN positions p ON l.purchase_id = p.purchase_id AND l.lot_number = p.lot_number
LEFT JOIN params prm ON p.purchase_id = prm.purchase_id 
    AND p.lot_number = prm.lot_number 
    AND p.position_number = prm.position_number
WHERE prm.param_name LIKE '%диагонал%'
ORDER BY l.purchase_id, l.lot_number, p.position_number;
        """)
        
        # Краткий итог
        print(f"\nПарсинг ноутбуков v2.0 завершен!")
        print(f"Найдено ноутбуков: {result}")
        print(f"Обработано закупок: {parser_instance.stats.requests_count}")
        print(f"Позиций: {parser_instance.stats.positions_total}")
        print(f"Параметров: {parser_instance.stats.params_total}")
        
    except KeyboardInterrupt:
        print("\nПрервано пользователем")
        interrupted = True
        return 1
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}")
        print(f"Критическая ошибка: {e}")
        return 1
    
    print("Программа завершена")
    return 0

if __name__ == "__main__":
    sys.exit(main())
