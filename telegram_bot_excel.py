# syllabus bot для ранепа-лицея
# читает данные из syllabuses.xlsx (файл должен уже существовать)

import telebot
from telebot import types
import openpyxl
import os
from dotenv import load_dotenv

# грузим токен из .env
load_dotenv()
TOKEN = os.getenv('BOT_TOKEN')
if not TOKEN:
    print("✗ нет BOT_TOKEN в файле .env")
    exit(1)

bot = telebot.TeleBot(TOKEN)
XLS_FILE = 'syllabuses.xlsx'

# храним контекст: chat_id → grade (строка) ИЛИ dict для выбора учителя
user_ctx = {}


def load_data():
    """читает силлабусы из готового экселя"""
    if not os.path.exists(XLS_FILE):
        print(f"✗ файл {XLS_FILE} не найден!")
        print("Запусти сначала: python update_syllabuses.py")
        exit(1)

    data = {}
    wb = openpyxl.load_workbook(XLS_FILE)

    for sheet in wb.sheetnames:
        cls = sheet.replace('Класс ', '')
        data[cls] = {}
        ws = wb[sheet]

        for row in ws.iter_rows(min_row=2, values_only=False):
            if len(row) < 3:
                continue
            subj, teach, link = row[0].value, row[1].value, row[2].value
            if not all([subj, teach, link]):
                continue

            if subj not in data[cls]:
                data[cls][subj] = []
            data[cls][subj].append({'teacher': teach, 'url': link})

    wb.close()
    print(f"✓ загружено {sum(len(s) for c in data.values() for s in c.values())} силлабусов")
    return data


# === клавиатуры ===
def main_kb():
    mk = types.ReplyKeyboardMarkup(resize_keyboard=True)
    mk.add('8 класс', '9 класс')
    mk.add('10 класс', '11 класс')
    mk.add('❓ Справка', '🎓 Mini App')
    return mk


def subjects_kb(grade, db):
    mk = types.ReplyKeyboardMarkup(resize_keyboard=True)
    if grade in db:
        subs = sorted(db[grade].keys())
        for i in range(0, len(subs), 2):
            if i + 1 < len(subs):
                mk.add(subs[i], subs[i + 1])
            else:
                mk.add(subs[i])
    mk.add('⬅️ Назад')
    return mk


def teachers_inline(teachers):
    mk = types.InlineKeyboardMarkup()
    for i, t in enumerate(teachers):
        mk.add(types.InlineKeyboardButton(t['teacher'], callback_data=f"t_{i}"))
    return mk


# загружаем базу при старте
db = load_data()


# === команды ===
@bot.message_handler(commands=['start'])
def start(msg):
    user_ctx[msg.chat.id] = None
    bot.send_message(
        msg.chat.id,
        "👋 Привет! Я бот с силлабусами лицея РАНХиГС.\nВыбери класс 👇",
        reply_markup=main_kb()
    )


@bot.message_handler(commands=['help'])
def help_cmd(msg):
    bot.send_message(
        msg.chat.id,
        "📋 Как пользоваться:\n"
        "1. Нажми на класс (8–11)\n"
        "2. Выбери предмет\n"
        "3. Если учителей несколько — выбери нужного\n"
        "4. Получи ссылку на силлабус\n\n"
        "Команды:\n"
        "/start — меню\n"
        "/help — эта справка\n"
        "/reload — обновить данные из Excel\n"
        "/webapp — открыть мини-приложение\n"
        "/rep — перейти в репозиторий",
        reply_markup=main_kb()
    )


@bot.message_handler(commands=['reload'])
def reload_db(msg):
    global db
    try:
        db = load_data()
        bot.send_message(msg.chat.id, "✓ данные обновлены", reply_markup=main_kb())
    except Exception as e:
        bot.send_message(msg.chat.id, f"✗ ошибка: {e}", reply_markup=main_kb())


@bot.message_handler(commands=['webapp'])
def send_webapp(msg):
    url = "https://pmikhail2009.github.io/syllabus-miniapp/  "
    mk = types.InlineKeyboardMarkup()
    mk.add(types.InlineKeyboardButton("🎓 Открыть силлабусы", web_app=types.WebAppInfo(url=url)))
    bot.send_message(msg.chat.id, "📱 Мини-приложение:", reply_markup=mk)


@bot.message_handler(commands=['rep'])
def send_repo(msg):
    url = "https://github.com/pmikhail2009/syllabus-miniapp  "
    mk = types.InlineKeyboardMarkup()
    mk.add(types.InlineKeyboardButton("🔗 GitHub", url=url))
    bot.send_message(msg.chat.id, f"📦 Исходный код:\n{url}", reply_markup=mk)


# === кнопки меню: выбор класса ===
@bot.message_handler(func=lambda m: m.text in ['8 класс', '9 класс', '10 класс', '11 класс'])
def pick_grade(msg):
    cid = msg.chat.id
    grade = msg.text.split()[0]
    # ✅ перезаписываем любой предыдущий контекст (даже dict с учителями)
    user_ctx[cid] = grade

    if grade not in db:
        bot.send_message(cid, "✗ нет данных для этого класса", reply_markup=main_kb())
        return

    count = len(db[grade])
    bot.send_message(
        cid,
        f"📚 {grade} класс ({count} предметов)\nВыбери предмет:",
        reply_markup=subjects_kb(grade, db)
    )


# === кнопки меню: Справка и Mini App ===
@bot.message_handler(func=lambda m: m.text == '❓ Справка')
def btn_help(msg):
    help_cmd(msg)


@bot.message_handler(func=lambda m: m.text == '🎓 Mini App')
def btn_webapp(msg):
    send_webapp(msg)


# === кнопка Назад ===
@bot.message_handler(func=lambda m: m.text == '⬅️ Назад')
def go_back(msg):
    cid = msg.chat.id
    user_ctx[cid] = None
    bot.send_message(cid, "📌 Выбери класс:", reply_markup=main_kb())


# === ОБЩИЙ ОБРАБОТЧИК: выбор предмета (должен быть ПОСЛЕ специфичных!) ===
@bot.message_handler(func=lambda m: True)
def pick_subject(msg):
    cid = msg.chat.id
    subj = msg.text
    ctx = user_ctx.get(cid)

    # ✅ Извлекаем класс из контекста (поддерживаем и строку, и словарь)
    if isinstance(ctx, dict):
        grade = ctx.get('grade')
    elif isinstance(ctx, str):
        grade = ctx
    else:
        grade = None

    # ✅ Если нажата кнопка класса — передаём обработку в pick_grade
    if subj in ['8 класс', '9 класс', '10 класс', '11 класс']:
        pick_grade(msg)
        return

    # ✅ Обработка кнопки «Назад»
    if subj == '⬅️ Назад':
        user_ctx[cid] = None
        bot.send_message(cid, "📌 Выбери класс:", reply_markup=main_kb())
        return

    # Проверка валидности класса и предмета
    if not grade or grade not in db or subj not in db[grade]:
        bot.send_message(cid, "✗ используйте кнопки меню", reply_markup=main_kb())
        user_ctx[cid] = None  # сброс контекста при ошибке
        return

    # Предмет найден — обрабатываем выбор
    teachers = db[grade][subj]

    if len(teachers) == 1:
        t = teachers[0]
        bot.send_message(
            cid,
            f"✅ {subj}\n📍 {grade} класс\n👨‍🏫 {t['teacher']}\n\n🔗 {t['url']}",
            reply_markup=main_kb()
        )
        user_ctx[cid] = None
        return

    # Несколько учителей — сохраняем контекст и показываем inline-кнопки
    user_ctx[cid] = {'grade': grade, 'subj': subj, 'teachers': teachers}
    bot.send_message(
        cid,
        f"👨‍🏫 Выберите учителя по предмету '{subj}':",
        reply_markup=teachers_inline(teachers)
    )


# === callback: выбор учителя ===
@bot.callback_query_handler(func=lambda c: c.data.startswith('t_'))
def on_teacher_pick(call):
    cid = call.message.chat.id
    idx = int(call.data.split('_')[1])
    ctx = user_ctx.get(cid)

    if not ctx or not isinstance(ctx, dict):
        bot.answer_callback_query(call.id, "сессия истекла, начни заново")
        return

    grade, subj, teachers = ctx['grade'], ctx['subj'], ctx['teachers']
    if idx >= len(teachers):
        bot.answer_callback_query(call.id, "ошибка выбора")
        return

    t = teachers[idx]
    bot.delete_message(cid, call.message.message_id)
    bot.send_message(
        cid,
        f"✅ {subj}\n📍 {grade} класс\n👨‍🏫 {t['teacher']}\n\n🔗 {t['url']}",
        reply_markup=main_kb()
    )
    user_ctx[cid] = None
    bot.answer_callback_query(call.id)


# === запуск ===
if __name__ == '__main__':
    print(f"🤖 бот запущен | файл: {XLS_FILE} | классов: {len(db)}")
    print("команды: /start /help /reload /webapp /rep")
    try:
        bot.infinity_polling()
    except KeyboardInterrupt:
        print("\n⏹ остановлен")
    except Exception as e:
        print(f"✗ упал: {e}")