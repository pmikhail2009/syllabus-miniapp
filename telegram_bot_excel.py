# ВАЖНАЯ ИНФА
# ПРО ИИ:
# Я использовал ИИ для 5 вещей
# 1) Проверка, оптимизация и исправление ошибок
# 2) Сделать приписки типа
# 3) Сделать classes_data Я крайне ленивый так, что создание того списка я поручил ИИ
# 4) Помощь с кнопками (я чуть не сдох пока пытался понять как их сделать ಠ_ಠ)
# 5) Я захотел сделать дизайн какой-то для excel таблички и поручил это тоже ИИ
# ¯\_(ツ)_/¯

import telebot
from telebot import types
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
from datetime import datetime

# ============================================
# ИНИЦИАЛИЗАЦИЯ БОТА
# ============================================
bot = telebot.TeleBot('8297243093:AAGK9lm0zqZ6g8WOkPAiMr9_SQ-PgAUN6I4')
EXCEL_FILE = 'syllabuses.xlsx'

# Словарь для хранения выбранного класса каждого пользователя
user_selected_grade = {}

# ============================================
# ФУНКЦИИ ДЛЯ РАБОТЫ С EXCEL
# ============================================
def create_excel_file():
    """Создание Excel файла с силлабусами по классам"""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Данные для каждого класса (с учителями)
    classes_data = {
        '8': [
            ('Русский язык', 'Е. О. Борисова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_rus_b_beo_2025.pdf'),
            ('Литература', 'Ю. Ю. Ишутин', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_lit_b_iyy_2025.pdf'),
            ('История', 'И. Ю. Романов', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_ist_b_riy_2025.pdf'),
            ('Обществознание', 'А. Н. Егорова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_obsh_b_ean_2025.pdf'),
            ('Право', 'А. В. Немова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_pravo_b_nav_2025.pdf'),
            ('Английский язык', 'Д. В. Плотникова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_eng_u_pdv_2025.pdf'),
            ('Английский язык', 'Е. В. Петрова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_eng_u_pev_2025.pdf'),
            ('Английский язык', 'У. Райс', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_eng_u_ru_2025.pdf'),
            ('Испанский язык', 'М. М. Шустовских', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_isp_b_1_shmm_2025.pdf'),
            ('Китайский язык', 'Н. Н. Шаталова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_kit_b_shnn_2025.pdf'),
            ('География', 'Е. В. Бетехтин', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_geo_b_bev_2025.pdf'),
            ('Экономика', 'И. А. Нимерницкая', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_econ_b_nia_2025.pdf'),
            ('Биология', 'В. В. Майчак', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_bio_b_mvv_2025.pdf'),
            ('Физика', 'А. А. Кочегаров', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_fiz_b_kaa_2025.pdf'),
            ('Химия', 'Д. В. Ефанов', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_him_b_edv_2025.pdf'),
            ('Алгебра', 'А. А. Коваленко', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_alg_b_kaa_2025.pdf'),
            ('Геометрия', 'А. А. Коваленко', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_geom_b_kaa_2025.pdf'),
            ('Теория вероятности', 'А. А. Коваленко', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_teorver_b_kaa_2025.pdf'),
            ('Информатика', 'Д.А. Панфилова, А.Ю. Величко', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_inf_b_pda_vay_2025.pdf'),
            ('Технология', 'Д. А. Панфилова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_8_1_tekn_b_pda_2025.pdf'),
        ],
        '9': [
            ('Русский язык', 'Е. О. Борисова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_rus_b_beo_2025.pdf'),
            ('Литература', 'Ю. Ю. Ишутин', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_lit_b_iyy_2025.pdf'),
            ('История', 'И. Ю. Романов', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_ist_b_riy_2025.pdf'),
            ('Обществознание', 'А. Н. Егорова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_obsh_b_ean_2025.pdf'),
            ('Право', 'А. В. Немова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_pravo_b_nav_2025.pdf'),
            ('Теория познания', 'С. М. Гнездилов', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_tok_b_gsm_2025.pdf'),
            ('Английский язык', 'У. Райс (1)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_eng_u_ru_2025.pdf'),
            ('Английский язык', 'У. Райс (2)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_eng_u_2_ru_2025.pdf'),
            ('Английский язык', 'Д. В. Плотникова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_eng_u_pdv_2025.pdf'),
            ('Английский язык', 'В. В. Хитрук (GWB1)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_eng_u_GWB1_hvv_2025.pdf'),
            ('Английский язык', 'В. В. Хитрук (GWB1+)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_eng_u_GWB1+_hvv_2025.pdf'),
            ('Английский язык', 'Е. Петрова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_eng_u_pev_2025.pdf'),
            ('Немецкий язык', 'О. Б. Темежникова (начинающие)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_nem_n_tob_2025.pdf'),
            ('Немецкий язык', 'О. Б. Темежникова (продолжающие-1)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_nem_p1_tob_2025.pdf'),
            ('Немецкий язык', 'О. Б. Темежникова (продолжающие-2)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_nem_p2_tob_2025.pdf'),
            ('Испанский язык', 'М. М. Шустовских (1)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_isp_b_1_shmm_2025.pdf'),
            ('Испанский язык', 'М. М. Шустовских (2)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_isp_b_2_shmm_2025.pdf'),
            ('География', 'Е. В. Бетехтин', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_geo_b_bev_2025.pdf'),
            ('Экономика', 'И. А. Нимерницкая', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_econ_b_nia_2025.pdf'),
            ('Биология', 'В. В. Майчак', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_bio_b_mvv_2025.pdf'),
            ('Физика', 'В. Н. Белянин', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_fiz_b_bvn_2025.pdf'),
            ('Химия', 'Д. В. Ефанов', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_him_b_edv_2025.pdf'),
            ('Алгебра', 'А. В. Пчелинцева', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_alg_b_pav_2025.pdf'),
            ('Геометрия', 'А. В. Пчелинцева', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_geom_b_pav_2025.pdf'),
            ('Теория вероятности', 'А. В. Пчелинцева', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_teorver_b_pav_2025.pdf'),
            ('Информатика', 'Д.А. Панфилова, А.Ю. Величко', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_inf_b_pda_vay_2025.pdf'),
            ('Технология', 'Д. А. Панфилова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_9_1_tekn_b_pda_2025.pdf'),
        ],
        '10': [
            ('Русский язык', 'Е. А. Пантелеева', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_rus_b_pea_2025.pdf'),
            ('Литература', 'П. Ф. Подковыркин (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_lit_b_ppf_2025.pdf'),
            ('Литература', 'П. Ф. Подковыркин (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_lit_u_ppf_2025.pdf'),
            ('История', 'А. Н. Кадыков', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_ist_b_kan_2025.pdf'),
            ('История', 'В. К. Герасимов', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_ist_u_gvk_2025.pdf'),
            ('Обществознание', 'С. М. Гнездилов', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_obsh_b_gsm_2025.pdf'),
            ('Обществознание', 'А. Н. Егорова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_obsh_b_ean_2025.pdf'),
            ('Право', 'А. В. Юсупов', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_prav_u_uav_2025.pdf'),
            ('Право', 'С.Н. Мореева', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_prav_u_msn_2025.pdf'),
            ('Теория и история искусства', 'Д. С. Матюнина', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_teor_isk_b_vkn_2025.pdf'),
            ('Теория познания', 'И. С. Курилович', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_ToK_u_kis_2025.pdf'),
            ('Английский язык', 'С. С. Московская', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_eng_u_mss_2025.pdf'),
            ('Английский язык', 'Е. В. Петрова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_eng_u_pev_2025.pdf'),
            ('Английский язык', 'Д. В. Плотникова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_eng_u_pdv_2025.pdf'),
            ('Немецкий язык', 'О. Б. Темежникова (начинающие)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_nem_n_tob_2025.pdf'),
            ('Немецкий язык', 'О. Б. Темежникова (продолжающие-1)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_nem_p1_tob_2025.pdf'),
            ('Немецкий язык', 'О. Б. Темежникова (продолжающие-2)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_nem_p2_tob_2025.pdf'),
            ('Немецкий язык', 'О. Б. Темежникова (продолжающие-3)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_nem_p3_tob_2025.pdf'),
            ('Испанский язык', 'М. М. Шустовских (1)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_isp_b_1_shmm_2025.pdf'),
            ('Испанский язык', 'М. М. Шустовских (2)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_isp_b_2_shmm_2025.pdf'),
            ('Испанский язык', 'М. М. Шустовских (3)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_isp_b_3_shmm_2025.pdf'),
            ('Испанский язык', 'М. М. Шустовских (4)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_isp_b_4_shmm_2025.pdf'),
            ('Итальянский язык', 'Е. С. Гурьева', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_ital_b_ges_2025.pdf'),
            ('Китайский язык', 'Н. Н. Шаталова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_kit_b_shnn_2025.pdf'),
            ('Арабский язык', 'Е. И. Смирнова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_arab_b_sei_2025.pdf'),
            ('География', 'Д. В. Степаньков', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_geo_b_sdv_2025.pdf'),
            ('Экономика', 'И. А. Нимерницкая (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_econ_b_nia_2025.pdf'),
            ('Экономика', 'И. А. Нимерницкая (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_econ_u_nia_2025.pdf'),
            ('Экономика', 'М. И. Мозгачёв', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_econ_u_mmi_2025.pdf'),
            ('Биология', 'В. В. Майчак (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_bio_b_mvv_2025.pdf'),
            ('Биология', 'В. В. Майчак (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_bio_u_mvv_2025.pdf'),
            ('Химия', 'Д. В. Ефанов', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_him_b_edv_2025.pdf'),
            ('Физика', 'А. А. Кочегаров', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_fiz_b_kaa_2025.pdf'),
            ('Алгебра', 'А. М. Городнов (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_alg_b_gam_2025.pdf'),
            ('Алгебра', 'В. К. Ушаков (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_alg_u_uvk_2025.pdf'),
            ('Алгебра', 'А. Ю. Величко (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_alg_u_vay_2025.pdf'),
            ('Алгебра', 'А. А. Коваленко (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_alg_b_kaa_2025.pdf'),
            ('Геометрия', 'А. М. Городнов (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_geom_b_gam_2025.pdf'),
            ('Геометрия', 'В. К. Ушаков (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_geom_u_uvk_2025.pdf'),
            ('Геометрия', 'А. Ю. Величко (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_geom_u_vay_2025.pdf'),
            ('Геометрия', 'А. А. Коваленко (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_geom_b_kaa_2025.pdf'),
            ('Информатика', 'Д. А. Панфилова, А. Ю. Величко', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_10_1_inf_u_pda_vay_2025.pdf'),
        ],
        '11': [
            ('Русский язык', 'Е. В. Гаранина, Е. О. Борисова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_rus_b_gev_beo_2025.pdf'),
            ('Литература', 'А. С. Сенилова (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_lit_b_sas_2025.pdf'),
            ('Литература', 'Е. В. Гаранина, Ю. Ю. Ишутин (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_lit_u_gev_iyy_2025.pdf'),
            ('Теория и история искусства', 'Д. С. Матюнина', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_teor_isk_b_mds_2025.pdf'),
            ('История', 'А. Н. Кадыков (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_ist_b_kan_2025.pdf'),
            ('История', 'А. Ф. Цветкова (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_ist_u_caf_2025.pdf'),
            ('Всеобщая политическая история', 'А. А. Атаманенко', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_polist_u_aaa_2025.pdf'),
            ('Политология', 'С. А. Кожеуров', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_polit_u_ksa_2025.pdf'),
            ('Право', 'А. В. Юсупов', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_prav_u_uav_2025.pdf'),
            ('Право', 'С.Н. Мореева (юридический)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_prav_u_2_msn_2025.pdf'),
            ('Обществознание', 'А. Н. Егорова (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_obsh_b_ean_2025.pdf'),
            ('Обществознание', 'А. Н. Егорова (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_obsh_u_ean_2025.pdf'),
            ('Социология', 'М. А. Серкина', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_soc_u_sma_2025.pdf'),
            ('Английский язык', 'В. В. Хитрук (Gateway В2)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_eng_u_GWB2_hvv_2025.pdf'),
            ('Английский язык', 'В. В. Хитрук (Gateway С1)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_eng_u_GWC1_hvv_2025.pdf'),
            ('Английский язык', 'В. В. Хитрук (Prepare 7)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_eng_u_PR7_hvv_2025.pdf'),
            ('Английский язык', 'С. С. Московская', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_eng_u_mss_2025.pdf'),
            ('Английский язык', 'А. Л. Ромашкина', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_eng_u_ral_2025.pdf'),
            ('Английский язык', 'Е. С. Гурина', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_eng_u_1_ges_2025.pdf'),
            ('Английский язык', 'У. Райс', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_eng_u_ru_2025.pdf'),
            ('Английский язык State Exam', 'Е. В. Петрова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_state_u_pev_2025.pdf'),
            ('Английский язык State Exam', 'С. С. Московская', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_state_u_mss_2025.pdf'),
            ('Немецкий язык', 'О. Б. Темежникова (продолжающие-1)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_nem_p1_tob_2025.pdf'),
            ('Немецкий язык', 'О. Б. Темежникова (продолжающие-2)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_nem_p2_tob_2025.pdf'),
            ('Немецкий язык', 'О. Б. Темежникова (продолжающие-3)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_nem_p3_tob_2025.pdf'),
            ('Испанский язык', 'Е. В. Львова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_isp_b_lev_2025.pdf'),
            ('Испанский язык', 'М. М. Шустовских (3)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_isp_b_3_shmm_2025.pdf'),
            ('Испанский язык', 'М. М. Шустовских (4)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_isp_b_4_shmm_2025.pdf'),
            ('Итальянский язык', 'Е. С. Гурьева', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_ital_b_ges_2025.pdf'),
            ('Французский язык', 'М. С. Салкина', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_fr_b_sms_2025.pdf'),
            ('Китайский язык', 'Н. Н. Шаталова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_kit_u_shnn_2025.pdf'),
            ('Арабский язык', 'Е. И. Смирнова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_arab_b_sei_2025.pdf'),
            ('География', 'Д. В. Степаньков', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_geo_b_sdv_2025.pdf'),
            ('Экономика', 'И. А. Нимерницкая (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_econ_b_nia_2025.pdf'),
            ('Экономика', 'И. А. Нимерницкая (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_econ_u_nia_2025.pdf'),
            ('Биология', 'В. В. Майчак', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_bio_u_mvv_2025.pdf'),
            ('Алгебра', 'В. К. Ушаков', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_alg_u_uvk_2025.pdf'),
            ('Алгебра', 'В. С. Осипова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_alg_u_ovs_2025.pdf'),
            ('Алгебра', 'А. В. Петров', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_alg_u_pav_2025.pdf'),
            ('Геометрия', 'В. К. Ушаков', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_geom_u_uvk_2025.pdf'),
            ('Геометрия', 'В. С. Осипова', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_geom_u_ovs_2025.pdf'),
            ('Геометрия', 'А. В. Петров', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_geom_u_pav_2025.pdf'),
            ('Информатика', 'Ю. П. Мартемьянов (базовый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_inf_b_myp_2025.pdf'),
            ('Информатика', 'Д. А. Панфилова, А. Ю. Величко (углублённый)', 'https://ranepa-lyceum.ru/docs/sil_25-26_1/sil_11_1_inf_u_pda_vay_2025.pdf'),
        ],
    }

    # Создание листов для каждого класса
    for grade, subjects_list in classes_data.items():
        ws = wb.create_sheet(title=f'Класс {grade}')

        # Заголовки (3 колонки)
        ws['A1'] = 'Предмет'
        ws['B1'] = 'Учитель'
        ws['C1'] = 'Ссылка на силлабус'

        for cell in ['A1', 'B1', 'C1']:
            ws[cell].fill = header_fill
            ws[cell].font = header_font
            ws[cell].border = border

        # Установка ширины колонок
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 25
        ws.column_dimensions['C'].width = 50

        # Добавление данных
        row = 2
        for subject, teacher, url in subjects_list:
            ws[f'A{row}'] = subject
            ws[f'B{row}'] = teacher
            ws[f'C{row}'] = url

            for cell in [f'A{row}', f'B{row}', f'C{row}']:
                ws[cell].border = border
                ws[cell].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

            row += 1

        # Высота строк
        ws.row_dimensions[1].height = 25
        for i in range(2, row):
            ws.row_dimensions[i].height = 30

    wb.save(EXCEL_FILE)
    print(f"✅ Excel файл '{EXCEL_FILE}' успешно создан!")


def load_syllabuses_from_excel():
    """Загрузка силлабусов из Excel файла"""
    if not os.path.exists(EXCEL_FILE):
        print(f"⚠️ Файл '{EXCEL_FILE}' не найден. Создаём новый...")
        create_excel_file()
    
    syllabuses = {}
    wb = openpyxl.load_workbook(EXCEL_FILE)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        grade = sheet_name.replace('Класс ', '')
        syllabuses[grade] = {}

        for row in ws.iter_rows(min_row=2, values_only=False):
            if len(row) < 3:
                continue

            subject_cell = row[0]
            teacher_cell = row[1]
            url_cell = row[2]

            if subject_cell.value and teacher_cell.value and url_cell.value:
                subject = subject_cell.value
                teacher = teacher_cell.value
                url = url_cell.value

                if subject not in syllabuses[grade]:
                    syllabuses[grade][subject] = []

                syllabuses[grade][subject].append({
                    'teacher': teacher,
                    'url': url
                })

    wb.close()
    return syllabuses


# ============================================
# КЛАВИАТУРЫ
# ============================================
def get_main_keyboard():
    """Главное меню с выбором класса"""
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    markup.add(types.KeyboardButton('8 класс'), types.KeyboardButton('9 класс'))
    markup.add(types.KeyboardButton('10 класс'), types.KeyboardButton('11 класс'))
    markup.add(types.KeyboardButton('❓ Справка'), types.KeyboardButton('🎓 Mini App'))
    return markup


def get_subjects_keyboard(grade, syllabuses):
    """Клавиатура с предметами для выбранного класса"""
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    if grade in syllabuses:
        subjects = sorted(list(syllabuses[grade].keys()))
        for i in range(0, len(subjects), 2):
            if i + 1 < len(subjects):
                markup.add(subjects[i], subjects[i + 1])
            else:
                markup.add(subjects[i])

    markup.add(types.KeyboardButton('⬅️ Назад'))
    return markup


def get_teachers_keyboard(teachers_list):
    """Клавиатура с выбором учителя (Inline кнопки)"""
    markup = types.InlineKeyboardMarkup()
    for idx, teacher_data in enumerate(teachers_list):
        button = types.InlineKeyboardButton(
            text=teacher_data['teacher'],
            callback_data=f"teacher_{idx}"
        )
        markup.add(button)
    return markup


# ============================================
# ЗАГРУЗКА ДАННЫХ ПРИ ЗАПУСКЕ
# ============================================
syllabuses = load_syllabuses_from_excel()


# ============================================
# ОБРАБОТЧИКИ КОМАНД
# ============================================
@bot.message_handler(commands=['start'])
def send_welcome(message):
    """Обработка команды /start"""
    chat_id = message.chat.id
    user_selected_grade[chat_id] = None
    welcome_text = (
        "👋 Добро пожаловать в бот силлабусов RANEPA Lyceum!\n\n"
        "Я помогу вам найти силлабус по интересующему вас предмету.\n\n"
        "📌 Выберите класс для начала:"
    )
    bot.send_message(chat_id, welcome_text, reply_markup=get_main_keyboard())


@bot.message_handler(commands=['help'])
def send_help(message):
    """Обработка команды /help"""
    chat_id = message.chat.id
    help_text = (
        "📖 Справка по боту\n\n"
        "Как пользоваться:\n"
        "1️⃣ Нажмите на класс (8–11)\n"
        "2️⃣ Выберите интересующий вас предмет\n"
        "3️⃣ Если у предмета несколько учителей - выберите нужного\n"
        "4️⃣ Получите ссылку на силлабус\n\n"
        "Команды:\n"
        "/start - Главное меню\n"
        "/help - Справка\n"
        "/reload - Перезагрузить данные из Excel\n"
        "/webapp - Открыть мини-приложение\n"
        "/rep - Перейти в репозиторий на GitHub\n\n"
        "На каждом экране есть кнопка '⬅️ Назад' для возврата в предыдущее меню."
    )
    bot.send_message(chat_id, help_text, parse_mode='HTML', reply_markup=get_main_keyboard())


@bot.message_handler(commands=['reload'])
def reload_data(message):
    """Перезагрузка данных из Excel файла"""
    chat_id = message.chat.id
    global syllabuses
    try:
        syllabuses = load_syllabuses_from_excel()
        bot.send_message(
            chat_id,
            "✅ Данные успешно перезагружены из файла Excel!",
            reply_markup=get_main_keyboard()
        )
    except Exception as e:
        bot.send_message(
            chat_id,
            f"❌ Ошибка при перезагрузке: {str(e)}",
            reply_markup=get_main_keyboard()
        )


@bot.message_handler(commands=['webapp'])
def send_miniapp(message):
    """Отправка кнопки для запуска Mini App"""
    chat_id = message.chat.id
    webapp_url = "https://pmikhail2009.github.io/syllabus-miniapp/"
    
    markup = types.InlineKeyboardMarkup()
    btn = types.InlineKeyboardButton(
        text="🎓 Открыть силлабусы",
        web_app=types.WebAppInfo(url=webapp_url)
    )
    markup.add(btn)
    
    bot.send_message(
        chat_id,
        "📱 Нажмите на кнопку ниже, чтобы открыть мини-приложение:",
        reply_markup=markup
    )


@bot.message_handler(commands=['rep'])
def send_repo_link(message):
    """Отправка ссылки на репозиторий"""
    chat_id = message.chat.id
    repo_url = "https://github.com/pmikhail2009/syllabus-miniapp"
    
    markup = types.InlineKeyboardMarkup()
    btn = types.InlineKeyboardButton(
        text="🔗 Открыть репозиторий",
        url=repo_url
    )
    markup.add(btn)
    
    bot.send_message(
        chat_id,
        f"📦 Исходный код проекта:\n{repo_url}",
        reply_markup=markup
    )


# ============================================
# ОБРАБОТЧИКИ СООБЩЕНИЙ
# ============================================
@bot.message_handler(func=lambda message: message.text in ['8 класс', '9 класс', '10 класс', '11 класс'])
def select_grade(message):
    """Обработка выбора класса"""
    chat_id = message.chat.id
    grade = message.text.replace(' класс', '')
    user_selected_grade[chat_id] = grade

    grade_text = {
        '8': '8 класс',
        '9': '9 класс',
        '10': '10 класс',
        '11': '11 класс',
    }

    if grade in syllabuses:
        subject_count = len(syllabuses[grade])
        response = f"📚 {grade_text[grade]} ({subject_count} предметов)\n\nВыберите предмет:"
        bot.send_message(
            chat_id,
            response,
            parse_mode='HTML',
            reply_markup=get_subjects_keyboard(grade, syllabuses)
        )
    else:
        bot.send_message(
            chat_id,
            f"❌ Данные для класса {grade} не найдены",
            reply_markup=get_main_keyboard()
        )


@bot.message_handler(func=lambda message: message.text == '❓ Справка')
def help_button(message):
    """Справка через кнопку"""
    send_help(message)


@bot.message_handler(func=lambda message: message.text == '🎓 Mini App')
def miniapp_button(message):
    """Кнопка Mini App в главном меню"""
    send_miniapp(message)


@bot.message_handler(func=lambda message: message.text == '⬅️ Назад')
def go_back(message):
    """Обработка кнопки 'Назад'"""
    chat_id = message.chat.id
    user_selected_grade[chat_id] = None
    bot.send_message(chat_id, "📌 Выберите класс:", reply_markup=get_main_keyboard())


@bot.message_handler(func=lambda message: True)
def select_subject(message):
    """Обработка выбора предмета"""
    chat_id = message.chat.id
    subject = message.text
    selected_grade = user_selected_grade.get(chat_id)

    if selected_grade and selected_grade in syllabuses:
        if subject in syllabuses[selected_grade]:
            teachers_list = syllabuses[selected_grade][subject]

            # Если учитель один - сразу выдаём силлабус
            if len(teachers_list) == 1:
                teacher_data = teachers_list[0]
                response = (
                    f"✅ {subject}\n\n"
                    f"📍 Класс: {selected_grade}\n"
                    f"👨‍🏫 Учитель: {teacher_data['teacher']}\n\n"
                    f"🔗 Откройте силлабус:\n{teacher_data['url']}"
                )
                bot.send_message(chat_id, response, parse_mode='HTML', reply_markup=get_main_keyboard())
                user_selected_grade[chat_id] = None
                return

            # Если учителей несколько - показываем выбор
            else:
                user_selected_grade[chat_id] = {
                    'grade': selected_grade,
                    'subject': subject,
                    'teachers': teachers_list
                }

                response = f"👨‍🏫 Выберите учителя по предмету '{subject}':"
                bot.send_message(
                    chat_id,
                    response,
                    reply_markup=get_teachers_keyboard(teachers_list)
                )
                return

    bot.send_message(
        chat_id,
        "❌ Предмет не найден. Используйте кнопки меню для навигации.",
        reply_markup=get_main_keyboard()
    )


# ============================================
# ОБРАБОТЧИКИ CALLBACK
# ============================================
@bot.callback_query_handler(func=lambda call: call.data.startswith('teacher_'))
def handle_teacher_selection(call):
    """Обработка выбора учителя через Inline кнопки"""
    chat_id = call.message.chat.id
    teacher_idx = int(call.data.split('_')[1])
    user_data = user_selected_grade.get(chat_id)

    if user_data and isinstance(user_data, dict):
        grade = user_data['grade']
        subject = user_data['subject']
        teachers_list = user_data['teachers']

        if teacher_idx < len(teachers_list):
            teacher_data = teachers_list[teacher_idx]

            response = (
                f"✅ {subject}\n\n"
                f"📍 Класс: {grade}\n"
                f"👨‍🏫 Учитель: {teacher_data['teacher']}\n\n"
                f"🔗 Откройте силлабус:\n{teacher_data['url']}"
            )

            # Удаляем сообщение с кнопками
            bot.delete_message(chat_id, call.message.message_id)

            # Отправляем результат
            bot.send_message(chat_id, response, parse_mode='HTML', reply_markup=get_main_keyboard())

            # Сбрасываем состояние
            user_selected_grade[chat_id] = None

            # Отвечаем на callback
            bot.answer_callback_query(call.id)


# ============================================
# ЗАПУСК БОТА
# ============================================
if __name__ == '__main__':
    print("🤖 Telegram бот с Excel поддержкой запущен...")
    print(f"📊 Excel файл: {EXCEL_FILE}")
    print(f"✅ Загружено предметов из {len(syllabuses)} классов")
    print(f"📱 Mini App: /webapp")
    print(f"🔗 Репозиторий: /rep")
    
    try:
        bot.infinity_polling()
    except KeyboardInterrupt:
        print("\n⏹️ Бот остановлен пользователем")
    except Exception as e:
        print(f"❌ Ошибка: {e}")