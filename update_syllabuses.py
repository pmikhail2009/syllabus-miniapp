#!/usr/bin/env python3
# update_syllabuses.py
# Запуск: python update_syllabuses.py

import sys
from scraper.parser import fetch_links
from scraper.excel_builder import create_or_update_excel
from config import EXCEL_PATH  # ← ДОБАВИТЬ ЭТУ СТРОКУ


def main():
    print("🔄 Парсинг силлабусов с ranepa-lyceum.ru...")
    
    links = fetch_links()
    if not links:
        print("✗ ничего не найдено")
        print("Возможные причины:")
        print("  • сайт изменил структуру")
        print("  • проблема с интернетом")
        print("  • сайт блокирует запросы")
        return 1
    
    print("📝 Создаю Excel-таблицу...")
    create_or_update_excel(links)
    
    # считаем статистику
    grades = set(item['grade'] for item in links)
    unknown = sum(1 for item in links if item['teacher'] == 'Требует уточнения')
    
    print(f"\n✓ Готово!")
    print(f"  • классов: {len(grades)} ({', '.join(sorted(grades))})")
    print(f"  • всего силлабусов: {len(links)}")
    print(f"  • учителей не распознано: {unknown}")
    
    if unknown > 0:
        print(f"\n⚠️  Открой {EXCEL_PATH} и проверь колонку 'Учитель'")
        print("   Там где 'Требует уточнения' — впиши ФИО вручную")
    
    return 0


if __name__ == '__main__':
    sys.exit(main())