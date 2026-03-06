# scraper/parser.py
# Парсит страницу с силлабусами и вытаскивает данные из имён файлов
# Логика: sil_{класс}_{группа}_{предмет}_{уровень}_{учитель}_{год}.pdf

import re
import requests
from bs4 import BeautifulSoup
from config import SUBJECTS, TEACHERS, LEVELS, BASE_URL, SYLLABUS_PAGE


def fetch_links():
    """Скачивает страницу и возвращает список ссылок на PDF"""
    try:
        resp = requests.get(SYLLABUS_PAGE, timeout=30, headers={
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        })
        resp.raise_for_status()
    except Exception as e:
        print(f"✗ не скачал страницу: {e}")
        return []

    soup = BeautifulSoup(resp.text, 'html.parser')
    links = []

    for a in soup.find_all('a', href=True):
        href = a['href']
        if href.endswith('.pdf') and 'sil' in href.lower():
            full_url = href if href.startswith('http') else BASE_URL + href
            info = parse_filename(href.split('/')[-1])
            if info:
                info['url'] = full_url
                links.append(info)

    print(f"✓ найдено силлабусов: {len(links)}")
    return links


def parse_filename(filename):
    """
    Парсит имя файла по шаблону:
    sil_{class}_{group}_{subject}_{level}_{teacher}_{year}.pdf
    
    Пример: sil_10_1_lit_u_ppf_2025.pdf
    → класс: 10, предмет: Литература, уровень: углублённый, учитель: П. Ф. Подковыркин
    """
    # убираем расширение
    name = filename.replace('.pdf', '').lower()
    parts = name.split('_')
    
    if len(parts) < 6 or parts[0] != 'sil':
        return None
    
    # 1) Класс — первое число после sil
    try:
        grade = parts[1]
        if grade not in ['8', '9', '10', '11']:
            return None
    except:
        return None
    
    # 2) Уровень — ищем b/u/n/p в соответствующей позиции
    level_code = parts[4] if len(parts) > 4 else 'b'
    level = LEVELS.get(level_code, 'базовый')
    
    # 3) Предмет — код в позиции 3
    subject_code = parts[3] if len(parts) > 3 else ''
    subject = SUBJECTS.get(subject_code, None)
    if not subject:
        # пробуем найти код в любой части имени
        for code, name in SUBJECTS.items():
            if code in name:
                subject = name
                break
    if not subject:
        subject = 'Неизвестный предмет'  # заглушка
    
    # 4) Учитель — код в позиции 5
    teacher_code = parts[5] if len(parts) > 5 else ''
    teacher = TEACHERS.get(teacher_code, None)
    
    # спец. обработка: если в учителе запятая (два препода)
    if not teacher and teacher_code in ['pda_vay', 'gev_beo', 'ge_iyy']:
        # составные коды
        teachers = []
        for code in teacher_code.split('_'):
            if code in TEACHERS:
                teachers.append(TEACHERS[code])
        teacher = ', '.join(teachers) if teachers else 'Требует уточнения'
    
    if not teacher:
        teacher = 'Требует уточнения'
    
    # 5) Доп. информация в скобках (например "базовый", "начинающие")
    # если в оригинальном имени файла было что-то в скобках — добавляем
    extra = ''
    if '(' in filename:
        match = re.search(r'\(([^)]+)\)', filename)
        if match:
            extra = match.group(1).strip()
    
    # формируем итоговое название предмета с уровнем
    if level != 'базовый' and extra:
        subject_display = f"{subject} ({extra})"
    elif level != 'базовый':
        subject_display = f"{subject} ({level})"
    elif extra:
        subject_display = f"{subject} ({extra})"
    else:
        subject_display = subject
    
    return {
        'grade': grade,
        'subject': subject_display,
        'subject_code': subject_code,
        'teacher': teacher,
        'level': level,
        'filename': filename
    }