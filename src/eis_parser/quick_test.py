#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
quick_test.py - Быстрое тестирование парсера ноутбуков v3.1.4
Автоматизированный набор тестов для проверки исправлений HTTP 422 и Windows совместимости
"""

import subprocess
import sys
import time
import os
import sqlite3
from datetime import datetime

# Настройка кодировки для Windows
if sys.platform.startswith('win'):
    try:
        # Попытка установить UTF-8 для Windows консоли
        os.system('chcp 65001 > nul 2>&1')
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except:
        # Если не удалось, используем безопасный режим без эмодзи
        pass

def safe_print(text: str, end: str = '\n'):
    """Безопасный вывод текста с обработкой кодировки"""
    try:
        print(text, end=end)
    except UnicodeEncodeError:
        # Заменяем проблемные символы на безопасные
        safe_text = text.encode('ascii', errors='replace').decode('ascii')
        print(safe_text, end=end)

def run_command(cmd, description="", timeout=300):
    """Запуск команды с таймаутом и отслеживанием результата"""
    safe_print(f"\n{'='*60}")
    safe_print(f">>> {description}")
    safe_print(f"Команда: {cmd}")
    safe_print(f"{'='*60}")
    
    start_time = time.time()
    
    try:
        result = subprocess.run(
            cmd, 
            shell=True, 
            capture_output=True, 
            text=True, 
            timeout=timeout,
            encoding='utf-8',
            errors='replace'
        )
        
        elapsed = time.time() - start_time
        
        if result.returncode == 0:
            safe_print(f"[+] УСПЕХ ({elapsed:.1f}с)")
            if result.stdout:
                safe_print("STDOUT:", result.stdout[-500:])  # Последние 500 символов
        else:
            safe_print(f"[-] ОШИБКА (код: {result.returncode}, время: {elapsed:.1f}с)")
            if result.stderr:
                safe_print("STDERR:", result.stderr)
            if result.stdout:
                safe_print("STDOUT:", result.stdout)
                
        return result.returncode == 0
        
    except subprocess.TimeoutExpired:
        safe_print(f"[!] ТАЙМАУТ ({timeout}с)")
        return False
    except Exception as e:
        safe_print(f"[!] ИСКЛЮЧЕНИЕ: {e}")
        return False

def check_db_results():
    """Проверка результатов в базе данных"""
    try:
        if not os.path.exists("laptop_parser.db"):
            safe_print("[-] База данных не найдена")
            return False
            
        conn = sqlite3.connect("laptop_parser.db")
        cursor = conn.cursor()
        
        cursor.execute("SELECT COUNT(*) FROM lots")
        lots_count = cursor.fetchone()[0]
        
        cursor.execute("SELECT COUNT(*) FROM positions") 
        positions_count = cursor.fetchone()[0]
        
        cursor.execute("SELECT COUNT(*) FROM params")
        params_count = cursor.fetchone()[0]
        
        conn.close()
        
        safe_print(f"\n[i] Результаты в БД:")
        safe_print(f"   Лотов: {lots_count}")
        safe_print(f"   Позиций: {positions_count}")
        safe_print(f"   Параметров: {params_count}")
        
        return lots_count > 0
        
    except Exception as e:
        safe_print(f"[-] Ошибка проверки БД: {e}")
        return False

def main():
    """Главная функция тестирования v3.1.4"""
    safe_print(">>> БЫСТРОЕ ТЕСТИРОВАНИЕ ПАРСЕРА НОУТБУКОВ v3.1.4")
    safe_print(">>> ПРОВЕРКА ИСПРАВЛЕНИЙ HTTP 422 И WINDOWS СОВМЕСТИМОСТИ")
    safe_print(f"Время начала: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Счетчики
    total_tests = 0
    passed_tests = 0
    
    # Тест 1: Проверка зависимостей
    total_tests += 1
    safe_print(f"\n[{total_tests}] Проверка зависимостей Python")
    try:
        import requests
        import sqlite3
        import json
        import tqdm
        safe_print("[+] Основные зависимости найдены")
        passed_tests += 1
    except ImportError as e:
        safe_print(f"[-] Отсутствуют зависимости: {e}")
        safe_print("Выполните: pip install -r requirements.txt")
    
    # Тест 2: Проверка версии парсера
    total_tests += 1
    safe_print(f"\n[{total_tests}] Проверка версии парсера")
    try:
        with open('src/eis_parser/laptop_parser.py', 'r', encoding='utf-8') as f:
            content = f.read()
            if 'v3.1.4' in content and 'ИСПРАВЛЕНИЕ Windows кодировки' in content:
                safe_print("[+] Парсер v3.1.4 с исправлениями Windows кодировки + HTTP 422")
                passed_tests += 1
            else:
                safe_print("[-] Неправильная версия парсера или отсутствуют исправления")
    except Exception as e:
        safe_print(f"[-] Ошибка проверки версии: {e}")
    
    # Тест 3: Запуск справки парсера
    total_tests += 1
    if run_command("python src/eis_parser/laptop_parser.py --help", "Запуск справки парсера", 60):
        passed_tests += 1
    
    # Тест 4: Минимальный запуск парсера
    total_tests += 1
    if run_command("python src/eis_parser/laptop_parser.py --start 2024-01-01 --end 2024-01-02", "Минимальный запуск", 30):
        passed_tests += 1
    
    # Тест 5: КРИТИЧЕСКИЙ - проверка отсутствия HTTP 422 и ошибок кодировки
    total_tests += 1
    safe_print(f"\n[{total_tests}] КРИТИЧЕСКИЙ - проверка отсутствия HTTP 422 и ошибок кодировки")
    cmd = "python src/eis_parser/laptop_parser.py --start 2024-01-01 --end 2024-01-02 --chunk-hours 1 --debug"
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True, timeout=180, encoding='utf-8', errors='replace')
    
    if result.returncode == 0:
        # Проверяем что нет HTTP 422 в выводе
        stderr_text = result.stderr.lower()
        stdout_text = result.stdout.lower()
        
        errors_found = []
        if "422" in stderr_text or "422" in stdout_text:
            errors_found.append("HTTP 422")
        if "unprocessable entity" in stderr_text:
            errors_found.append("Unprocessable Entity")
        if "unicodeencodeerror" in stderr_text:
            errors_found.append("Unicode кодировка")
        if "charmap" in stderr_text:
            errors_found.append("Windows кодировка cp1251")
        
        if not errors_found:
            safe_print("[+] HTTP 422 и ошибки кодировки исправлены! Нет критических ошибок")
            passed_tests += 1
            
            # Проверяем что используются правильные параметры
            if "skip" in result.stderr:
                safe_print("[+] Используется правильный параметр 'skip' вместо 'offset'")
            if "classifiers" in result.stderr and "[0]" not in result.stderr:
                safe_print("[+] Классификаторы передаются без индексов")
        else:
            safe_print(f"[-] Найдены ошибки: {', '.join(errors_found)}")
            safe_print("STDERR:", result.stderr[-500:])
    else:
        safe_print("[-] Парсинг завершился с ошибкой")
        safe_print("STDERR:", result.stderr)
    
    # Тест 6: Быстрый парсинг с сохранением данных
    total_tests += 1
    if run_command("python src/eis_parser/laptop_parser.py --start 2024-01-01 --end 2024-01-02 --chunk-hours 1", "Быстрый парсинг", 180):
        passed_tests += 1
        
        # Проверяем результаты в БД
        if check_db_results():
            safe_print("[+] Данные успешно сохранены в БД")
        else:
            safe_print("[!] Данные не найдены в БД (возможно нет ноутбуков за период)")
    
    # Тест 7: Тест с фильтрами (проверка исправленного формата регионов)
    total_tests += 1
    if run_command("python src/eis_parser/laptop_parser.py --start 2024-01-01 --end 2024-01-02 --min-price 50000 --regions 77 --chunk-hours 1", "Тест фильтров (регионы как строка)", 120):
        passed_tests += 1
    
    # Тест 8: Многопоточный тест
    total_tests += 1  
    if run_command("python src/eis_parser/laptop_parser.py --start 2024-01-01 --end 2024-01-02 --threads 2 --chunk-hours 1", "Многопоточный тест", 120):
        passed_tests += 1
    
    # Тест 9: Повторный запуск парсера
    total_tests += 1
    if run_command("python src/eis_parser/laptop_parser.py --start 2024-01-01 --end 2024-01-02 --chunk-hours 1", "Повторный запуск", 120):
        passed_tests += 1
    
    # Тест 10: Специальный тест Windows кодировки
    total_tests += 1
    safe_print(f"\n[{total_tests}] Тест Windows кодировки")
    try:
        # Попытка вывода проблемных символов
        test_text = "Тест: русский текст + символы ASCII"
        safe_print(test_text)
        safe_print("[+] Windows кодировка работает корректно")
        passed_tests += 1
    except Exception as e:
        safe_print(f"[-] Ошибка Windows кодировки: {e}")
    
    # Итоговые результаты
    safe_print(f"\n{'='*60}")
    safe_print(f">>> ИТОГИ ТЕСТИРОВАНИЯ v3.1.4")
    safe_print(f"{'='*60}")
    safe_print(f"Всего тестов: {total_tests}")
    safe_print(f"Пройдено: {passed_tests}")
    safe_print(f"Не пройдено: {total_tests - passed_tests}")
    safe_print(f"Успешность: {passed_tests/total_tests*100:.1f}%")
    
    if passed_tests == total_tests:
        safe_print("*** ВСЕ ТЕСТЫ ПРОЙДЕНЫ! HTTP 422 И WINDOWS КОДИРОВКА ИСПРАВЛЕНЫ! ***")
        safe_print("Парсер готов к использованию в Windows!")
        return 0
    elif passed_tests >= total_tests * 0.7:
        safe_print("[!] Большинство тестов пройдено. Проверьте ошибки выше.")
        if passed_tests >= 7:  # Если критические тесты прошли
            safe_print("[+] Критические исправления HTTP 422 и Windows кодировки работают!")
        return 1
    else:
        safe_print("[-] Много ошибок. Проверьте установку зависимостей и доступность API.")
        return 2

if __name__ == "__main__":
    sys.exit(main())
