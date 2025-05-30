#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
parser.py – Универсальный парсер закупок v11.0
Поддерживает любые категории товаров через YAML-конфиг

КЛЮЧЕВЫЕ ВОЗМОЖНОСТИ:
✅ Поддержка множества категорий товаров через categories.yml
✅ Универсальная система фильтрации (ключевые слова + параметры)
✅ CLI с расширенными опциями и конфиг-override
✅ Многопоточность с автоопределением CPU
✅ HTTP кэширование с requests-cache (TTL 24ч)
✅ Адаптивный rate-limiting по типу сервера
✅ Детальная статистика по категориям
✅ Экспорт с группировкой по категориям
✅ Расширенная схема БД с поддержкой категорий
"""

# ============================================================================
# SECTION 0: IMPORTS & GLOBALS
# ============================================================================

import os
import sys
import json
import time
import csv
import sqlite3
import logging
import argparse
import shelve
import hashlib
import signal
import threading
import yaml
from datetime import datetime, timedelta, date
from typing import Dict, List, Optional, Tuple, Any, Union
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from dataclasses import dataclass, asdict

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

try:
    import aiohttp
    import asyncio
    ASYNC_AVAILABLE = True
except ImportError:
    ASYNC_AVAILABLE = False

# Базовые константы API
BASE_URL = "https://v2.gosplan.info/fz44"
LIST_ENDPOINT = "/purchases"
HEADERS = {"Accept": "application/json", "Accept-Encoding": "gzip"}

# Лимиты API по типу сервера
API_LIMITS = {
    "v2.gosplan.info": 600,  # продуктовый
    "test.gosplan.info": 10,  # тестовый
}

DEFAULT_PAGE_LIMIT = 100
MAX_PAGINATION_OFFSET = 10_000
MAX_THREADS = min(2 * os.cpu_count(), 20)
EXPORT_FILENAME_MASK = "%Y%m%d"

# Глобальные переменные
interrupted = False
logger = None

def setup_logging(debug: bool = False):
    """Настройка логирования с файлом parser_%Y%m%d_%H%M.log"""
    global logger
    
    log_filename = datetime.now().strftime("parser_%Y%m%d_%H%M.log")
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
    logger.info(f"Универсальный парсер закупок v11.0 запущен. Лог: {log_filename}")

def signal_handler(signum, frame):
    """Обработчик прерывания"""
    global interrupted
    interrupted = True
    print("\n⚠️ Получен сигнал прерывания...")
    if logger:
        logger.info("Получен сигнал прерывания")

signal.signal(signal.SIGINT, signal_handler)
signal.signal(signal.SIGTERM, signal_handler)

# ============================================================================
# SECTION 1: CONFIG LOADER (YAML + CLI MERGE)
# ============================================================================

@dataclass
class CategoryConfig:
    """Конфигурация категории товаров"""
    name: str
    okpd_prefix: Optional[str] = None
    ktru: Optional[List[str]] = None
    include_keywords: Optional[List[str]] = None
    exclude_keywords: Optional[List[str]] = None
    required_params: Optional[List[str]] = None
    min_price: Optional[float] = None
    max_price: Optional[float] = None
    regions: Optional[List[int]] = None

class ConfigLoader:
    """Загрузчик конфигурации с объединением YAML + CLI"""
    
    def __init__(self, config_path: str = "categories.yml"):
        self.config_path = config_path
        self.categories = {}
        self.load_yaml_config()
    
    def load_yaml_config(self):
        """Загрузка конфигурации из YAML"""
        if not os.path.exists(self.config_path):
            # Создаем пример конфига если его нет
            self.create_example_config()
        
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                config_data = yaml.safe_load(f) or {}
            
            for category_name, category_data in config_data.items():
                self.categories[category_name] = CategoryConfig(
                    name=category_data.get('name', category_name),
                    okpd_prefix=category_data.get('okpd_prefix'),
                    ktru=category_data.get('ktru', []),
                    include_keywords=category_data.get('include_keywords', []),
                    exclude_keywords=category_data.get('exclude_keywords', []),
                    required_params=category_data.get('required_params', []),
                    min_price=category_data.get('min_price'),
                    max_price=category_data.get('max_price'),
                    regions=category_data.get('regions')
                )
            
            logger.info(f"Загружено {len(self.categories)} категорий из {self.config_path}")
            
        except Exception as e:
            logger.error(f"Ошибка загрузки конфига {self.config_path}: {e}")
            # Fallback на встроенные категории
            self.create_default_categories()
    
    def create_example_config(self):
        """Создание примера конфигурации"""
        example_config = {
            'laptop': {
                'name': 'Ноутбук',
                'okpd_prefix': '26.20.11.110',
                'ktru': [
                    '26.20.11.110-00000138',
                    '26.20.11.110-00000139',
                    '26.20.11.110-00000140',
                    '26.20.11.110-00000141'
                ],
                'include_keywords': ['ноутбук', 'ультрабук', 'laptop', 'портативный компьютер'],
                'exclude_keywords': ['планшет', 'системный блок', 'сервер'],
                'required_params': ['размер диагонали экрана'],
                'min_price': 10000,
                'regions': []
            },
            'monitor': {
                'name': 'Монитор',
                'okpd_prefix': '26.20.17.110', 
                'ktru': [
                    '26.20.17.110-00000001',
                    '26.20.17.110-00000002'
                ],
                'include_keywords': ['монитор', 'дисплей', 'экран'],
                'exclude_keywords': ['ноутбук', 'планшет'],
                'required_params': ['размер диагонали', 'разрешение'],
                'min_price': 5000,
                'regions': []
            },
            'server': {
                'name': 'Сервер',
                'okpd_prefix': '26.20.14.000',
                'ktru': [
                    '26.20.14.000-00000001',
                    '26.20.14.000-00000188'
                ],
                'include_keywords': ['сервер', 'server'],
                'exclude_keywords': ['серверная стойка', 'серверный шкаф'],
                'required_params': ['количество установленных процессоров'],
                'min_price': 50000,
                'regions': []
            }
        }
        
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                yaml.dump(example_config, f, default_flow_style=False, 
                         allow_unicode=True, sort_keys=False)
            print(f"📄 Создан пример конфига: {self.config_path}")
        except Exception as e:
            logger.error(f"Ошибка создания примера конфига: {e}")
    
    def create_default_categories(self):
        """Создание категорий по умолчанию в памяти"""
        self.categories = {
            'laptop': CategoryConfig(
                name='laptop',
                ktru=['26.20.11.110-00000138', '26.20.11.110-00000139'],
                include_keywords=['ноутбук', 'ультрабук', 'laptop'],
                exclude_keywords=['планшет', 'системный блок'],
                required_params=['размер диагонали экрана']
            )
        }
        logger.warning("Использованы категории по умолчанию")
    
    def get_category(self, category_name: str) -> Optional[CategoryConfig]:
        """Получение конфигурации категории"""
        return self.categories.get(category_name)
    
    def get_all_categories(self) -> Dict[str, CategoryConfig]:
        """Получение всех категорий"""
        return self.categories
    
    def merge_with_cli(self, category_config: CategoryConfig, cli_args) -> CategoryConfig:
        """Объединение конфига с CLI параметрами (CLI перекрывает конфиг)"""
        merged = CategoryConfig(
            name=category_config.name,
            okpd_prefix=category_config.okpd_prefix,
            ktru=category_config.ktru,
            include_keywords=category_config.include_keywords,
            exclude_keywords=category_config.exclude_keywords,
            required_params=category_config.required_params,
            min_price=cli_args.min_price if cli_args.min_price is not None else category_config.min_price,
            max_price=cli_args.max_price if cli_args.max_price is not None else category_config.max_price,
            regions=cli_args.regions if cli_args.regions is not None else category_config.regions
        )
        return merged

# ============================================================================
# SECTION 2: DATA CLASSES
# ============================================================================

@dataclass
class RequestStats:
    """Статистика запросов с разбивкой по категориям"""
    requests_count: int = 0
    pages_fetched: int = 0
    lots_total: int = 0
    lots_by_category: Dict[str, int] = None
    errors_count: int = 0
    start_time: float = 0
    
    def __post_init__(self):
        if self.lots_by_category is None:
            self.lots_by_category = {}

# ============================================================================
# SECTION 3: HTTP CLIENT (С КЭШЕМ)
# ============================================================================

class AdaptiveRateLimiter:
    """Адаптивный rate limiter с учетом типа сервера"""
    
    def __init__(self, base_url: str):
        self.latencies = []
        self.current_delay = 0.5
        self.lock = threading.Lock()
        
        # Определяем лимит по URL сервера
        for url_pattern, limit in API_LIMITS.items():
            if url_pattern in base_url:
                self.requests_per_minute = limit
                break
        else:
            self.requests_per_minute = 600  # по умолчанию
        
        self.window_start = time.time()
        self.requests_in_window = 0
        
        logger.info(f"Rate limiter: {self.requests_per_minute} запросов/мин для {base_url}")
    
    def record_request(self, latency: float):
        """Записать запрос и латентность"""
        with self.lock:
            current_time = time.time()
            
            # Сброс окна каждые 60 секунд
            if current_time - self.window_start >= 60:
                self.window_start = current_time
                self.requests_in_window = 0
            
            self.requests_in_window += 1
            self.latencies.append(latency)
            
            if len(self.latencies) > 100:
                self.latencies.pop(0)
    
    def get_delay(self) -> float:
        """Получить задержку с учетом лимитов и латентности"""
        with self.lock:
            # Базовая задержка для соблюдения rate limit
            if self.requests_in_window >= self.requests_per_minute:
                return 60.0  # Ждем до следующего окна
            
            # Адаптивная задержка на основе латентности
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
    """HTTP клиент с кэшированием и ретраями"""
    
    def __init__(self, cache_ttl_hours: int = 24):
        self.base_url = BASE_URL
        self.session = self._create_session()
        self.rate_limiter = AdaptiveRateLimiter(self.base_url)
        
        # Настройка кэша если доступен
        if CACHE_AVAILABLE:
            requests_cache.install_cache(
                cache_name='http_cache',
                backend='sqlite',
                expire_after=cache_ttl_hours * 3600
            )
            logger.info(f"HTTP кэш включен: TTL {cache_ttl_hours}ч")
        else:
            logger.warning("requests-cache недоступен, кэширование отключено")
    
    def _create_session(self):
        """Создание HTTP сессии"""
        session = requests.Session()
        
        # Ретраи для 5xx ошибок (429 обрабатываем отдельно)
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
        """GET запрос с rate limiting и ретраями"""
        global interrupted
        
        if interrupted:
            return None
        
        url = f"{self.base_url}{endpoint}"
        max_retries = 5
        
        for attempt in range(max_retries):
            start_time = time.monotonic()
            
            try:
                # Применяем rate limiting
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
                    # HTTP 429: экспоненциальная задержка
                    delay = min(60 * (2 ** attempt), 300)
                    logger.warning(f"HTTP 429, пауза {delay}s (попытка {attempt + 1})")
                    time.sleep(delay)
                    continue
                elif 500 <= resp.status_code < 600:
                    delay = min(2 ** attempt, 16)
                    logger.warning(f"HTTP {resp.status_code}, back-off {delay}s (попытка {attempt + 1})")
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

# ============================================================================
# SECTION 4: DB LAYER (РАСШИРЕНО)
# ============================================================================

class Database:
    """Управление базой данных с поддержкой категорий"""
    
    def __init__(self, db_path: str = "universal_parser.db"):
        self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        """Инициализация схемы БД с категориями"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Таблица категорий
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS category (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                description TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Таблица лотов с категориями
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS lots (
                purchase_id TEXT NOT NULL,
                lot_number TEXT NOT NULL,
                category_id TEXT NOT NULL,
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
                PRIMARY KEY (purchase_id, lot_number),
                FOREIGN KEY (category_id) REFERENCES category (id)
            )
        ''')
        
        # Таблица параметров
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS params (
                purchase_id TEXT NOT NULL,
                lot_number TEXT NOT NULL,
                param_name TEXT NOT NULL,
                param_value TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (purchase_id, lot_number) REFERENCES lots (purchase_id, lot_number)
            )
        ''')
        
        # Таблица для last_seen
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS kv_store (
                key TEXT PRIMARY KEY,
                value TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        # Индексы для максимальной производительности
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_lots_category ON lots(category_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_lots_purchase ON lots(purchase_id)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_lots_price ON lots(start_price_rub)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_lots_date ON lots(published_at)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_params_lot ON params(purchase_id, lot_number)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_params_name ON params(param_name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_params_value ON params(param_value)')
        
        # Полнотекстовый индекс для быстрого поиска по названиям
        try:
            cursor.execute('''
                CREATE VIRTUAL TABLE IF NOT EXISTS lots_fts USING fts5(
                    purchase_id, lot_name, customer_name,
                    tokenize='unicode61 remove_diacritics 1'
                )
            ''')
            logger.info("Полнотекстовый индекс FTS5 создан")
        except Exception as e:
            logger.warning(f"FTS5 недоступен: {e}")
        
        conn.commit()
        conn.close()
        
        logger.info(f"База данных инициализирована: {self.db_path}")
    
    def ensure_category(self, category_id: str, category_name: str) -> str:
        """Убедиться что категория существует, создать если нет"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('SELECT id FROM category WHERE id = ?', (category_id,))
        if not cursor.fetchone():
            cursor.execute('''
                INSERT INTO category (id, name, description)
                VALUES (?, ?, ?)
            ''', (category_id, category_name, f"Категория {category_name}"))
            logger.debug(f"Создана категория: {category_id} ({category_name})")
        
        conn.commit()
        conn.close()
        return category_id
    
    def get_last_seen(self, category_id: str = "default") -> Optional[str]:
        """Получить last_seen дату для категории"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('SELECT value FROM kv_store WHERE key = ?', (f'last_seen_{category_id}',))
        result = cursor.fetchone()
        conn.close()
        return result[0] if result else None
    
    def set_last_seen(self, date_str: str, category_id: str = "default"):
        """Установить last_seen дату для категории"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO kv_store (key, value)
            VALUES (?, ?)
        ''', (f'last_seen_{category_id}', date_str))
        conn.commit()
        conn.close()
        logger.debug(f"Обновлен last_seen для {category_id}: {date_str}")
    
    def save_lots(self, lots_data: List[Dict], category_id: str) -> int:
        """Сохранение лотов с привязкой к категории"""
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
                    (purchase_id, lot_number, category_id, lot_name, okpd2, ktru, qty, uom,
                     start_price_rub, date_bid_start, date_bid_end, delivery_deadline,
                     customer_name, customer_inn, published_at, contract_price, winner_name)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    lot_data.get('purchase_id', ''),
                    lot_data.get('lot_number', ''),
                    category_id,
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
                
                # Удаляем старые параметры
                cursor.execute('''
                    DELETE FROM params 
                    WHERE purchase_id = ? AND lot_number = ?
                ''', (lot_data.get('purchase_id', ''), lot_data.get('lot_number', '')))
                
                # Сохраняем параметры
                dynamic_params = lot_data.get('dynamic_params', {})
                for param_name, param_value in dynamic_params.items():
                    cursor.execute('''
                        INSERT INTO params (purchase_id, lot_number, param_name, param_value)
                        VALUES (?, ?, ?, ?)
                    ''', (
                        lot_data.get('purchase_id', ''),
                        lot_data.get('lot_number', ''),
                        param_name,
                        str(param_value)
                    ))
                
                saved_count += 1
                
            except Exception as e:
                logger.error(f"Ошибка сохранения лота {lot_data.get('purchase_id', '')}: {e}")
        
        conn.commit()
        conn.close()
        
        if saved_count > 0:
            logger.info(f"Сохранено {saved_count} лотов категории {category_id}")
        
        return saved_count

# ============================================================================
# SECTION 5: UNIVERSAL FILTER ENGINE
# ============================================================================

def normalize_text(text: str) -> str:
    """Нормализация текста"""
    if not text:
        return ""
    return " ".join(text.lower().strip().split())

def extract_number_from_string(text: str) -> Optional[float]:
    """Извлечение числа из строки"""
    import re
    if not text:
        return None
    
    match = re.search(r'(\d+(?:[.,]\d+)?)', str(text))
    if match:
        try:
            return float(match.group(1).replace(',', '.'))
        except ValueError:
            pass
    return None

def matches_category(lot: Dict, category_cfg: CategoryConfig) -> Tuple[bool, List[str]]:
    """
    Универсальная проверка соответствия лота категории
    
    Проверяет:
    1. Include/exclude keywords в названии и характеристиках
    2. Наличие всех required_params с непустыми значениями
    3. Числовые диапазоны (если заданы)
    
    Возвращает: (matches, reasons)
    """
    reasons = []
    
    # Собираем весь текст для поиска ключевых слов
    searchable_text = ""
    if lot.get('lot_name'):
        searchable_text += " " + str(lot['lot_name'])
    if lot.get('object_info'):
        searchable_text += " " + str(lot['object_info'])
    
    # Добавляем текст из характеристик
    def _extract_text_from_nested(data, path=""):
        nonlocal searchable_text
        if isinstance(data, dict):
            for key, value in data.items():
                if isinstance(value, str):
                    searchable_text += " " + value
                elif isinstance(value, (dict, list)):
                    _extract_text_from_nested(value, f"{path}.{key}")
        elif isinstance(data, list):
            for item in data:
                _extract_text_from_nested(item, path)
        elif isinstance(data, str):
            searchable_text += " " + data
    
    # Извлекаем текст из различных разделов
    for section in ['characteristics', 'specifications', 'parameters', 'items']:
        if section in lot:
            _extract_text_from_nested(lot[section], section)
    
    searchable_text = normalize_text(searchable_text)
    
    # 1. Проверка include keywords
    if category_cfg.include_keywords:
        found_include = False
        for keyword in category_cfg.include_keywords:
            if normalize_text(keyword) in searchable_text:
                found_include = True
                reasons.append(f"include_keyword:{keyword}")
                break
        
        if not found_include:
            return False, ["no_include_keywords"]
    
    # 2. Проверка exclude keywords
    if category_cfg.exclude_keywords:
        for keyword in category_cfg.exclude_keywords:
            if normalize_text(keyword) in searchable_text:
                return False, [f"exclude_keyword:{keyword}"]
    
    # 3. Проверка required_params
    if category_cfg.required_params:
        found_params = set()
        
        def _find_required_params(data, path=""):
            if isinstance(data, dict):
                # Ищем пары parameter/value
                if 'parameter' in data or 'param_name' in data or 'name' in data:
                    param_name = data.get('parameter') or data.get('param_name') or data.get('name', '')
                    param_value = data.get('value') or data.get('param_value') or data.get('val', '')
                    
                    if param_name and param_value:
                        normalized_param = normalize_text(str(param_name))
                        
                        # Проверяем каждый required параметр
                        for required_param in category_cfg.required_params:
                            normalized_required = normalize_text(required_param)
                            if normalized_required in normalized_param:
                                # Проверяем что значение не пустое
                                if param_value and str(param_value).strip():
                                    found_params.add(required_param)
                                    reasons.append(f"required_param:{required_param}={param_value}")
                
                # Рекурсивно ищем дальше
                for key, value in data.items():
                    _find_required_params(value, f"{path}.{key}")
            
            elif isinstance(data, (list, tuple)):
                for i, item in enumerate(data):
                    _find_required_params(item, f"{path}[{i}]")
        
        # Ищем параметры во всем лоте
        _find_required_params(lot)
        
        # Проверяем что найдены ВСЕ required параметры
        missing_params = set(category_cfg.required_params) - found_params
        if missing_params:
            return False, [f"missing_required_params:{list(missing_params)}"]
    
    # 4. Проверка ценовых диапазонов
    if category_cfg.min_price is not None or category_cfg.max_price is not None:
        price = lot.get('start_price_rub') or lot.get('max_price') or lot.get('price')
        if price is not None:
            try:
                price_float = float(price)
                if category_cfg.min_price is not None and price_float < category_cfg.min_price:
                    return False, [f"price_below_min:{price_float}<{category_cfg.min_price}"]
                if category_cfg.max_price is not None and price_float > category_cfg.max_price:
                    return False, [f"price_above_max:{price_float}>{category_cfg.max_price}"]
                reasons.append(f"price_ok:{price_float}")
            except (ValueError, TypeError):
                pass
    
    return True, reasons

# ============================================================================
# SECTION 6: PARSER ENGINE
# ============================================================================

def batched(iterable, n):
    """Разбивка списка на батчи"""
    import itertools
    iterator = iter(iterable)
    while True:
        batch = list(itertools.islice(iterator, n))
        if not batch:
            break
        yield batch

class UniversalParser:
    """Универсальный парсер закупок"""
    
    def __init__(self, config_loader: ConfigLoader, threads: int = MAX_THREADS, 
                 chunk_hours: int = 24, use_async: bool = False, show_progress: bool = False):
        self.config_loader = config_loader
        self.http_client = HTTPClient()
        self.database = Database()
        self.stats = RequestStats()
        self.threads = max(1, min(threads, MAX_THREADS))
        self.chunk_hours = min(chunk_hours, 24)
        self.use_async = use_async and ASYNC_AVAILABLE
        self.show_progress = show_progress and TQDM_AVAILABLE
        
        if self.use_async:
            logger.info("Асинхронный режим включен")
        
        logger.info(f"Парсер инициализирован (потоков: {self.threads}, chunk: {self.chunk_hours}ч)")
    
    def _create_time_chunks(self, start_date: datetime, end_date: datetime) -> List[Tuple[datetime, datetime]]:
        """Создание временных чанков для предотвращения переполнения пагинации"""
        chunks = []
        step = timedelta(hours=self.chunk_hours)
        
        curr_start = start_date
        while curr_start < end_date:
            curr_end = min(curr_start + step - timedelta(seconds=1), end_date)
            chunks.append((curr_start, curr_end))
            curr_start = curr_end + timedelta(seconds=1)
        
        logger.info(f"Создано {len(chunks)} временных чанков по {self.chunk_hours}ч")
        return chunks
    
    def _extract_lot_data(self, purchase_data: Dict, category_cfg: CategoryConfig) -> List[Dict]:
        """Извлечение данных лотов из закупки для категории"""
        extracted_lots = []
        
        # Базовая информация о закупке
        purchase_id = purchase_data.get('purchase_number', '')
        published_at = purchase_data.get('published_at', '')
        
        # Информация о заказчике
        customers = purchase_data.get('customers', [])
        customer_name = customers[0] if customers else ''
        customer_inn = customers[0] if customers else ''
        
        # Обрабатываем лоты
        lots = purchase_data.get('lots', [])
        if not lots and 'items' in purchase_data:
            # Fallback на items если нет lots
            lots = purchase_data['items']
        
        for lot_idx, lot in enumerate(lots):
            self.stats.lots_total += 1
            
            # Проверяем соответствие категории
            matches, match_reasons = matches_category(lot, category_cfg)
            
            if not matches:
                logger.debug(f"Лот {purchase_id}[{lot_idx}] не подходит для {category_cfg.name}: {match_reasons}")
                continue
            
            # Увеличиваем счетчик для категории
            if category_cfg.name not in self.stats.lots_by_category:
                self.stats.lots_by_category[category_cfg.name] = 0
            self.stats.lots_by_category[category_cfg.name] += 1
            
            # Извлекаем основные поля
            lot_data = {
                'purchase_id': purchase_id,
                'lot_number': str(lot.get('lotNumber', lot.get('lot_number', lot_idx))),
                'lot_name': lot.get('lotName', lot.get('name', '')),
                'okpd2': lot.get('okpd2', lot.get('okpd2Code', '')),
                'ktru': lot.get('ktru', lot.get('ktruCode', lot.get('code', ''))),
                'qty': self._extract_number(lot, ['quantity', 'qty', 'value']),
                'uom': lot.get('unitName', lot.get('unit', {}).get('shortName', '')),
                'start_price_rub': self._extract_number(purchase_data, ['max_price', 'startPrice']) or self._extract_number(lot, ['price']),
                'date_bid_start': purchase_data.get('startOfReception', purchase_data.get('collecting_started_at', '')),
                'date_bid_end': purchase_data.get('collecting_finished_at', purchase_data.get('endOfReception', '')),
                'delivery_deadline': lot.get('delivery_end_date', lot.get('deliveryDate', '')),
                'customer_name': customer_name,
                'customer_inn': customer_inn,
                'published_at': published_at,
                'contract_price': self._extract_number(purchase_data, ['contract_price', 'contractPrice']),
                'winner_name': lot.get('winner', {}).get('name', '') if isinstance(lot.get('winner'), dict) else '',
                'dynamic_params': self._extract_dynamic_params(lot)
            }
            
            extracted_lots.append(lot_data)
            
            logger.debug(f"✅ Лот {purchase_id}[{lot_idx}] соответствует {category_cfg.name}: {match_reasons}")
        
        return extracted_lots
    
    def _extract_number(self, data: Dict, possible_keys: List[str]) -> Optional[float]:
        """Извлечение числового значения"""
        for key in possible_keys:
            if key in data and data[key] is not None:
                try:
                    return float(data[key])
                except (ValueError, TypeError):
                    continue
        return None
    
    def _extract_dynamic_params(self, lot: Dict) -> Dict[str, str]:
        """Извлечение динамических параметров"""
        params = {}
        
        def _extract_params_recursive(data, prefix=""):
            if isinstance(data, dict):
                if 'parameter' in data and 'value' in data:
                    param_name = str(data['parameter'])
                    param_value = str(data['value'])
                    if param_name and param_value:
                        full_name = f"{prefix}{param_name}" if prefix else param_name
                        params[full_name] = param_value
                
                for key, value in data.items():
                    if key not in ['parameter', 'value']:
                        new_prefix = f"{prefix}{key}." if prefix else f"{key}."
                        _extract_params_recursive(value, new_prefix)
            
            elif isinstance(data, (list, tuple)):
                for i, item in enumerate(data):
                    _extract_params_recursive(item, f"{prefix}[{i}].")
        
        for section in ['characteristics', 'specifications', 'parameters', 'items']:
            if section in lot:
                _extract_params_recursive(lot[section], f"{section}.")
        
        return params
    
    def _fetch_purchases_for_category(self, category_cfg: CategoryConfig, 
                                    start_date: datetime, end_date: datetime) -> List[Dict]:
        """Получение закупок для категории с оптимизированным chunking"""
        
        # Формируем коды для поиска
        search_codes = []
        if category_cfg.ktru:
            search_codes.extend(category_cfg.ktru)
        if category_cfg.okpd_prefix:
            search_codes.append(category_cfg.okpd_prefix)
        
        if not search_codes:
            logger.warning(f"Нет кодов для поиска в категории {category_cfg.name}")
            return []
        
        logger.info(f"Поиск для категории {category_cfg.name}: {len(search_codes)} кодов")
        
        # Создаем временные чанки для предотвращения переполнения пагинации
        time_chunks = self._create_time_chunks(start_date, end_date)
        
        # Разбиваем коды на батчи по 10
        code_batches = list(batched(search_codes, 10))
        
        all_purchases = []
        total_combinations = len(time_chunks) * len(code_batches)
        
        logger.info(f"Обработка {len(time_chunks)} временных чанков × {len(code_batches)} батчей кодов = {total_combinations} комбинаций")
        
        # Создаем итератор для progress bar
        if self.show_progress:
            combinations_iter = tqdm(
                [(chunk, batch) for chunk in time_chunks for batch in code_batches],
                desc=f"{category_cfg.name}",
                unit="batch"
            )
        else:
            combinations_iter = [(chunk, batch) for chunk in time_chunks for batch in code_batches]
        
        def process_combination(chunk_and_batch):
            """Обработка одной комбинации временного чанка и батча кодов"""
            (chunk_start, chunk_end), code_batch = chunk_and_batch
            
            try:
                purchases = self._fetch_period_recursive(code_batch, chunk_start, chunk_end, category_cfg)
                
                # Теплое сохранение: извлекаем и сохраняем лоты сразу после получения батча
                if purchases:
                    batch_lots = []
                    for purchase in purchases:
                        lots = self._extract_lot_data(purchase, category_cfg)
                        batch_lots.extend(lots)
                    
                    if batch_lots:
                        # Сохраняем сразу, не накапливая в памяти
                        self.database.save_lots(batch_lots, category_cfg.name)
                        logger.debug(f"Теплое сохранение: {len(batch_lots)} лотов для {category_cfg.name}")
                
                return len(purchases)
                
            except Exception as e:
                logger.error(f"Ошибка обработки комбинации {chunk_start}-{chunk_end} + {code_batch[:2]}...: {e}")
                self.stats.errors_count += 1
                
                # Авто-retry на уровне батча при множественных ошибках
                if "5" in str(e) or "429" in str(e):
                    logger.warning("Множественные ошибки сервера, пауза 120 сек...")
                    time.sleep(120)
                    try:
                        purchases = self._fetch_period_recursive(code_batch, chunk_start, chunk_end, category_cfg)
                        return len(purchases)
                    except Exception as retry_e:
                        logger.error(f"Повторная ошибка: {retry_e}")
                
                return 0
        
        # Обрабатываем комбинации
        total_purchases = 0
        
        if self.threads == 1:
            # Последовательная обработка для отладки
            for combination in combinations_iter:
                if interrupted:
                    break
                count = process_combination(combination)
                total_purchases += count
        else:
            # Параллельная обработка
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
        
        logger.info(f"Загрузка для {category_cfg.name} завершена: {total_purchases} закупок")
        return []  # Возвращаем пустой список т.к. все уже сохранено через теплое сохранение
    
    def _fetch_period_recursive(self, code_batch: List[str], 
                              start_date: datetime, end_date: datetime,
                              category_cfg: CategoryConfig) -> List[Dict]:
        """Рекурсивная загрузка периода с разбиением и поддержкой регионов"""
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
            
            # Добавляем коды как classifiers[i]
            for i, code in enumerate(code_batch[:10]):
                params[f"classifiers[{i}]"] = code
            
            # Добавляем регионы если указаны в конфигурации
            if category_cfg.regions:
                for i, region in enumerate(category_cfg.regions):
                    params[f"region"] = region  # API поддерживает множественные region параметры
            
            self.stats.requests_count += 1
            page_data = self.http_client.get(LIST_ENDPOINT, params)
            
            if not page_data:
                break
            
            self.stats.pages_fetched += 1
            all_data.extend(page_data)
            
            if len(page_data) < DEFAULT_PAGE_LIMIT:
                break
            
            offset += DEFAULT_PAGE_LIMIT
            
            # Если достигли предела пагинации - разбиваем период
            if offset >= MAX_PAGINATION_OFFSET:
                logger.info(f"Достигнут предел пагинации, разбиваем период")
                
                total_seconds = (end_date - start_date).total_seconds()
                mid_date = start_date + timedelta(seconds=total_seconds / 2)
                
                first_half = self._fetch_period_recursive(code_batch, start_date, mid_date - timedelta(days=1), category_cfg)
                second_half = self._fetch_period_recursive(code_batch, mid_date, end_date, category_cfg)
                
                all_data.extend(first_half)
                all_data.extend(second_half)
                break
        
        return all_data
    
    def parse_categories(self, categories: List[str], start_date: datetime, end_date: datetime) -> Dict[str, int]:
        """Парсинг выбранных категорий с теплым сохранением"""
        self.stats.start_time = time.time()
        self.stats.lots_by_category = {}
        
        results = {}
        
        # Создаем progress bar для категорий если нужно
        if self.show_progress:
            categories_iter = tqdm(categories, desc="Категории", unit="cat")
        else:
            categories_iter = categories
        
        for category_name in categories_iter:
            if interrupted:
                break
            
            category_cfg = self.config_loader.get_category(category_name)
            if not category_cfg:
                logger.error(f"Категория {category_name} не найдена в конфигурации")
                continue
            
            # Убеждаемся что категория существует в БД
            self.database.ensure_category(category_name, category_cfg.name)
            
            logger.info(f"🔍 Начинаем поиск для категории: {category_cfg.name}")
            print(f"🔍 Обрабатываем категорию: {category_cfg.name}")
            
            # Получаем закупки (теперь сохранение происходит внутри _fetch_purchases_for_category)
            self._fetch_purchases_for_category(category_cfg, start_date, end_date)
            
            # Подсчитываем сохраненные лоты из БД
            conn = sqlite3.connect(self.database.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT COUNT(*) FROM lots WHERE category_id = ?', (category_name,))
            saved_count = cursor.fetchone()[0]
            conn.close()
            
            results[category_name] = saved_count
            if saved_count > 0:
                print(f"✅ {category_cfg.name}: найдено {saved_count} лотов")
            else:
                print(f"⚠️ {category_cfg.name}: лоты не найдены")
            
            # Обновляем last_seen для категории
            self.database.set_last_seen(end_date.strftime("%Y-%m-%dT%H:%M:%SZ"), category_name)
        
        return results

# ============================================================================
# SECTION 7: EXPORTERS
# ============================================================================

class DataExporter:
    """Экспорт данных с группировкой по категориям"""
    
    def __init__(self, database: Database):
        self.database = database
    
    def get_export_data(self, category_filter: Optional[str] = None) -> Tuple[List[Dict], List[str]]:
        """Получение данных для экспорта"""
        conn = sqlite3.connect(self.database.db_path)
        cursor = conn.cursor()
        
        # Формируем запрос с учетом фильтра категории
        if category_filter and category_filter != 'all':
            where_clause = "WHERE l.category_id = ?"
            query_params = [category_filter]
        else:
            where_clause = ""
            query_params = []
        
        # Получаем данные лотов с категориями и параметрами
        cursor.execute(f'''
            SELECT l.*, c.name as category_name,
                   GROUP_CONCAT(p.param_name || '=' || p.param_value, '|') as params_str
            FROM lots l
            LEFT JOIN category c ON l.category_id = c.id
            LEFT JOIN params p ON l.purchase_id = p.purchase_id AND l.lot_number = p.lot_number
            {where_clause}
            GROUP BY l.purchase_id, l.lot_number
            ORDER BY l.created_at DESC
        ''', query_params)
        
        results = cursor.fetchall()
        
        # Получаем все уникальные параметры
        if category_filter and category_filter != 'all':
            cursor.execute('''
                SELECT DISTINCT p.param_name 
                FROM params p
                JOIN lots l ON p.purchase_id = l.purchase_id AND p.lot_number = l.lot_number
                WHERE l.category_id = ?
                ORDER BY p.param_name
            ''', [category_filter])
        else:
            cursor.execute('SELECT DISTINCT param_name FROM params ORDER BY param_name')
        
        all_params = [row[0] for row in cursor.fetchall()]
        
        conn.close()
        
        # Формируем список словарей для экспорта
        export_data = []
        for row in results:
            # Базовые поля
            item = {
                'purchase_id': row[0],
                'lot_number': row[1],
                'category_id': row[2],
                'lot_name': row[3],
                'okpd2': row[4],
                'ktru': row[5],
                'qty': row[6],
                'uom': row[7],
                'start_price_rub': row[8],
                'date_bid_start': row[9],
                'date_bid_end': row[10],
                'delivery_deadline': row[11],
                'customer_name': row[12],
                'customer_inn': row[13],
                'published_at': row[14],
                'contract_price': row[15],
                'winner_name': row[16],
                'category_name': row[18] or row[2],  # name из category или id
            }
            
            # Парсим динамические параметры
            params_str = row[19]  # последнее поле с параметрами
            dynamic_params = {}
            if params_str:
                for param_pair in params_str.split('|'):
                    if '=' in param_pair:
                        param_name, param_value = param_pair.split('=', 1)
                        dynamic_params[param_name] = param_value
            
            # Добавляем динамические параметры
            for param_name in all_params:
                item[param_name] = dynamic_params.get(param_name, '')
            
            export_data.append(item)
        
        return export_data, all_params
    
    def export_to_csv(self, category_filter: str = 'all', filename: str = None) -> str:
        """Экспорт в CSV с группировкой по категориям"""
        if not filename:
            if category_filter == 'all':
                filename = f"all_{datetime.now().strftime(EXPORT_FILENAME_MASK)}.csv"
            else:
                filename = f"{category_filter}_{datetime.now().strftime(EXPORT_FILENAME_MASK)}.csv"
        
        data, dynamic_params = self.get_export_data(category_filter)
        
        if not data:
            print("⚠️ Нет данных для экспорта")
            return filename
        
        # Формируем fieldnames
        fixed_columns = [
            'purchase_id', 'lot_number', 'category_name', 'lot_name', 'okpd2', 'ktru',
            'qty', 'uom', 'start_price_rub',
            'date_bid_start', 'date_bid_end', 'delivery_deadline',
            'customer_name', 'customer_inn',
            'published_at', 'contract_price', 'winner_name'
        ]
        
        fieldnames = fixed_columns + sorted(dynamic_params)
        
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(data)
            
            print(f"📄 Экспорт CSV: {filename}")
            print(f"📊 Записей: {len(data)}")
            print(f"📋 Динамических параметров: {len(dynamic_params)}")
            
            # Статистика по категориям
            if category_filter == 'all':
                category_counts = {}
                for item in data:
                    cat = item.get('category_name', 'unknown')
                    category_counts[cat] = category_counts.get(cat, 0) + 1
                
                print("📊 По категориям:")
                for cat, count in sorted(category_counts.items()):
                    print(f"  {cat}: {count}")
            
            logger.info(f"CSV экспорт завершен: {filename}, записей: {len(data)}")
            
        except Exception as e:
            logger.error(f"Ошибка экспорта CSV: {e}")
            raise
        
        return filename
    
    def export_to_xlsx(self, category_filter: str = 'all', filename: str = None) -> str:
        """Экспорт в XLSX с автошириной и группировкой"""
        if not XLSX_AVAILABLE:
            raise RuntimeError("openpyxl не установлен. Установите: pip install openpyxl")
        
        if not filename:
            if category_filter == 'all':
                filename = f"all_{datetime.now().strftime(EXPORT_FILENAME_MASK)}.xlsx"
            else:
                filename = f"{category_filter}_{datetime.now().strftime(EXPORT_FILENAME_MASK)}.xlsx"
        
        data, dynamic_params = self.get_export_data(category_filter)
        
        if not data:
            print("⚠️ Нет данных для экспорта")
            return filename
        
        # Те же fieldnames что и для CSV
        fixed_columns = [
            'purchase_id', 'lot_number', 'category_name', 'lot_name', 'okpd2', 'ktru',
            'qty', 'uom', 'start_price_rub',
            'date_bid_start', 'date_bid_end', 'delivery_deadline',
            'customer_name', 'customer_inn',
            'published_at', 'contract_price', 'winner_name'
        ]
        
        fieldnames = fixed_columns + sorted(dynamic_params)
        
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Универсальный экспорт"
            
            # Заголовки с форматированием
            for col, header in enumerate(fieldnames, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
            
            # Данные
            for row_idx, item in enumerate(data, 2):
                for col_idx, field in enumerate(fieldnames, 1):
                    ws.cell(row=row_idx, column=col_idx, value=item.get(field, ''))
            
            # Автоширина колонок
            for column in ws.columns:
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
                ws.column_dimensions[column_letter].width = adjusted_width
            
            wb.save(filename)
            
            print(f"📄 Экспорт XLSX: {filename}")
            print(f"📊 Записей: {len(data)}")
            print(f"📋 Динамических параметров: {len(dynamic_params)}")
            
            logger.info(f"XLSX экспорт завершен: {filename}, записей: {len(data)}")
            
        except Exception as e:
            logger.error(f"Ошибка экспорта XLSX: {e}")
            raise
        
        return filename

# ============================================================================
# SECTION 8: MAIN() С CLI
# ============================================================================

def print_final_statistics(stats: RequestStats):
    """Печать финальной статистики по категориям"""
    runtime_seconds = time.time() - stats.start_time
    runtime_str = f"{int(runtime_seconds // 60)}m {int(runtime_seconds % 60)}s"
    
    print(f"\n" + "="*60)
    print(f"📊 ФИНАЛЬНАЯ СТАТИСТИКА УНИВЕРСАЛЬНОГО ПАРСЕРА v11.0")
    print(f"="*60)
    print(f"Requests        : {stats.requests_count}")
    print(f"Pages fetched   : {stats.pages_fetched}")
    print(f"Lots total      : {stats.lots_total}")
    
    # Статистика по категориям
    total_category_lots = sum(stats.lots_by_category.values())
    for category, count in sorted(stats.lots_by_category.items()):
        print(f"Lots {category:<12}: {count}")
    
    print(f"Errors          : {stats.errors_count}")
    print(f"Runtime         : {runtime_str}")
    print(f"="*60)

def main():
    """Главная функция с расширенным CLI"""
    global interrupted
    
    # CLI аргументы с новыми флагами для максимальной производительности
    parser = argparse.ArgumentParser(description="Универсальный парсер закупок v11.0")
    parser.add_argument('--category', type=str, default='all', 
                       help="Категория товаров или 'all' для всех")
    parser.add_argument('--config', type=str, default='categories.yml',
                       help="Путь к YAML конфигу")
    parser.add_argument('--start', type=str, help="Начало периода (YYYY-MM-DD)")
    parser.add_argument('--end', type=str, help="Конец периода (YYYY-MM-DD)")
    parser.add_argument('--min-price', type=float, help="Минимальная цена")
    parser.add_argument('--max-price', type=float, help="Максимальная цена")
    parser.add_argument('--outfile', type=str, help="Имя файла без расширения")
    parser.add_argument('--xlsx', action='store_true', help="Дополнительно экспорт в Excel")
    parser.add_argument('--threads', type=int, default=MAX_THREADS, 
                       help=f"Число потоков (1-{MAX_THREADS})")
    parser.add_argument('--debug', action='store_true', help="Включить DEBUG логирование")
    parser.add_argument('--chunk-hours', type=int, default=24,
                       help='Шаг разрезки периода, часы (<=24)')
    parser.add_argument('--async', dest='use_async', action='store_true',
                       help='Включить асинхронный режим (aiohttp)')
    parser.add_argument('--progress', action='store_true',
                       help='Показывать progress bar')
    parser.add_argument('--regions', type=int, nargs='*',
                       help='Коды регионов (77 78 для МСК+СПб)')
    
    args = parser.parse_args()
    
    # Настройка логирования
    setup_logging(args.debug)
    
    print("🚀 Универсальный парсер закупок v11.0 ТУРБО")
    print("✨ Максимальная производительность и надежность")
    print("⚡ Chunked периоды + теплое сохранение + FTS поиск")
    print("🎯 Progress bar + асинхронный режим + авто-retry")
    if not TQDM_AVAILABLE:
        print("⚠️ Для progress bar установите: pip install tqdm")
    if not ASYNC_AVAILABLE:
        print("⚠️ Для async режима установите: pip install aiohttp")
    print("")
    
    try:
        # Инициализация компонентов с новыми параметрами
        config_loader = ConfigLoader(args.config)
        database = Database()
        parser_instance = UniversalParser(
            config_loader, 
            threads=args.threads,
            chunk_hours=args.chunk_hours,
            use_async=args.use_async,
            show_progress=args.progress
        )
        exporter = DataExporter(database)
        
        # Определение категорий для обработки
        if args.category == 'all':
            categories_to_process = list(config_loader.get_all_categories().keys())
            print(f"📋 Будут обработаны все категории: {', '.join(categories_to_process)}")
        else:
            if args.category not in config_loader.get_all_categories():
                print(f"❌ Категория '{args.category}' не найдена в конфиге")
                available = ', '.join(config_loader.get_all_categories().keys())
                print(f"Доступные категории: {available}")
                return 1
            categories_to_process = [args.category]
            print(f"📋 Будет обработана категория: {args.category}")
        
        # Применение CLI-override к конфигам категорий
        if (args.min_price is not None or args.max_price is not None or 
            args.regions is not None):
            for category_name in categories_to_process:
                category_cfg = config_loader.get_category(category_name)
                merged_cfg = config_loader.merge_with_cli(category_cfg, args)
                config_loader.categories[category_name] = merged_cfg
                print(f"  📝 {category_name}: min_price={merged_cfg.min_price}, max_price={merged_cfg.max_price}, regions={merged_cfg.regions}")
        
        # Определение периода поиска
        if args.start:
            start_date = datetime.strptime(args.start, "%Y-%m-%d")
        else:
            # Автоматически используем last_seen или yesterday
            last_seen_dates = []
            for category in categories_to_process:
                last_seen_str = database.get_last_seen(category)
                if last_seen_str:
                    try:
                        last_seen_date = datetime.fromisoformat(last_seen_str.replace('Z', '+00:00'))
                        last_seen_dates.append(last_seen_date)
                    except:
                        pass
            
            if last_seen_dates:
                start_date = min(last_seen_dates)
                print(f"📅 Используем min(last_seen) как начальную дату: {start_date.date()}")
            else:
                start_date = datetime.now() - timedelta(days=1)
                print(f"📅 last_seen не найден, используем вчера: {start_date.date()}")
        
        if args.end:
            end_date = datetime.strptime(args.end, "%Y-%m-%d")
        else:
            end_date = datetime.now()
        
        # Поиск
        logger.info(f"Запуск поиска с {start_date} по {end_date} для категорий: {categories_to_process}")
        results = parser_instance.parse_categories(categories_to_process, start_date, end_date)
        
        # Экспорт
        export_category = args.category if args.category != 'all' else 'all'
        
        csv_filename = args.outfile if args.outfile else None
        csv_path = exporter.export_to_csv(export_category, csv_filename)
        
        # Дополнительный экспорт в XLSX если требуется
        if args.xlsx:
            if XLSX_AVAILABLE:
                xlsx_filename = args.outfile if args.outfile else None
                if xlsx_filename and not xlsx_filename.endswith('.xlsx'):
                    xlsx_filename = xlsx_filename.replace('.csv', '') + '.xlsx'
                exporter.export_to_xlsx(export_category, xlsx_filename)
            else:
                print("❌ openpyxl не установлен, XLSX экспорт пропущен")
        
        # Финальная статистика
        print_final_statistics(parser_instance.stats)
        
        # Краткий итог
        total_found = sum(results.values())
        print(f"\n🎉 Парсинг завершен!")
        print(f"📊 Всего найдено лотов: {total_found}")
        for category, count in results.items():
            print(f"  {category}: {count}")
        
    except KeyboardInterrupt:
        print("\n⚠️ Прервано пользователем")
        interrupted = True
        return 1
    except Exception as e:
        logger.error(f"Критическая ошибка: {e}")
        print(f"💥 Критическая ошибка: {e}")
        return 1
    
    print("🔚 Программа завершена")
    return 0

if __name__ == "__main__":
    sys.exit(main())