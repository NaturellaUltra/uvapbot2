import logging
import sqlite3
import datetime
from telegram import (
    Update, ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton
)
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler, filters,
    ContextTypes, ConversationHandler, CallbackQueryHandler
)
from openpyxl import Workbook
from io import BytesIO
from dotenv import load_dotenv
import os

load_dotenv()

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)

REGISTER_NAME, REGISTER_DEPARTMENT, WAIT_REASON = range(3)

# Администраторы 
ADMIN_IDS = {717329852,653756588}

#  Названия отделов 
DEPARTMENTS = [
    "Управление по выявлению административных правонарушений",
    "Отдел центральных районов",
    "Отдел северных районов",
    "Отдел южных районов",
    "Отдел правобережных районов",
    "Отдел координации"
]

departments_keyboard = [[KeyboardButton(dept)] for dept in DEPARTMENTS]
main_button = [[KeyboardButton("🚪 Сообщить об убытии с рабочего места")]]
admin_button = [[KeyboardButton("📊 Получить отчёт")]]

# База данных
conn = sqlite3.connect("bot_data.db", check_same_thread=False)
cursor = conn.cursor()
cursor.execute("""
CREATE TABLE IF NOT EXISTS users (
    user_id INTEGER PRIMARY KEY,
    full_name TEXT,
    department TEXT,
    is_admin INTEGER
)
""")
cursor.execute("""
CREATE TABLE IF NOT EXISTS departures (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    reason TEXT,
    timestamp DATETIME
)
""")
conn.commit()

def is_registered(user_id: int) -> bool:
    cursor.execute("SELECT 1 FROM users WHERE user_id = ?", (user_id,))
    return cursor.fetchone() is not None

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    print("CHAT ID:", update.effective_chat.id)
    user_id = update.effective_user.id
    if is_registered(user_id):
        cursor.execute("SELECT is_admin FROM users WHERE user_id = ?", (user_id,))
        is_admin = cursor.fetchone()[0]
        keyboard = main_button.copy()
        if is_admin:
            keyboard += admin_button
        await update.message.reply_text("✅ Вы уже зарегистрированы.", reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
        return ConversationHandler.END
    await update.message.reply_text("👋 Здравствуйте! Пожалуйста, введите ваше полное ФИО (например: Иванов Сергей Петрович):")
    return REGISTER_NAME

async def register_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    name = update.message.text.strip()
    if len(name.split()) < 3:
        await update.message.reply_text("⚠️ Пожалуйста, укажите ФИО полностью (Фамилия Имя Отчество).")
        return REGISTER_NAME
    context.user_data["full_name"] = name
    await update.message.reply_text("✅ Теперь выберите ваш отдел:", reply_markup=ReplyKeyboardMarkup(departments_keyboard, resize_keyboard=True))
    return REGISTER_DEPARTMENT

async def register_department(update: Update, context: ContextTypes.DEFAULT_TYPE):
    department = update.message.text
    user_id = update.effective_user.id
    full_name = context.user_data["full_name"]
    is_admin = 1 if user_id in ADMIN_IDS else 0

    cursor.execute("INSERT OR REPLACE INTO users (user_id, full_name, department, is_admin) VALUES (?, ?, ?, ?)",
                   (user_id, full_name, department, is_admin))
    conn.commit()

    keyboard = main_button.copy()
    if is_admin:
        keyboard += admin_button
        message = "🎉 Регистрация завершена. Вы добавлены в администраторы и Вам предоставлен доступ к выгрузке файла статистики."
    else:
        message = "🎉 Регистрация завершена."
    await update.message.reply_text(message, reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True))
    return ConversationHandler.END

async def handle_departure(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_registered(user_id):
        await update.message.reply_text("⚠️ Сначала необходимо пройти регистрацию. Напишите /start.")
        return ConversationHandler.END
    await update.message.reply_text("✏️ Напишите, куда вы уходите и во сколько планируете вернуться (например: «в МФЦ, вернусь в 14:30»):")
    return WAIT_REASON


async def save_departure(update: Update, context: ContextTypes.DEFAULT_TYPE):
    now = datetime.datetime.now()
    if now.weekday() >= 5 or not (9 <= now.hour < 18):
        await update.message.reply_text("⏰ Вне рабочего времени — бот завершает работу.")
        return ConversationHandler.END

    user_id = update.effective_user.id
    reason = update.message.text
    timestamp = datetime.datetime.now()

    # Сохраняем в базу данных
    cursor.execute("INSERT INTO departures (user_id, reason, timestamp) VALUES (?, ?, ?)", (user_id, reason, timestamp))
    conn.commit()

    # Получаем данные пользователя
    cursor.execute("SELECT full_name, department FROM users WHERE user_id = ?", (user_id,))
    full_name, department = cursor.fetchone()

    # Записываем в лог-файл
    with open("departures_log.txt", "a", encoding="utf-8") as f:
        f.write(f"[{timestamp.strftime('%Y-%m-%d %H:%M')}] {full_name} | {department} | {reason}\n")

    # 📣 Отправка уведомления в отдельный чат
    NOTIFY_CHAT_ID = -1002685248701  # ← замените на ваш настоящий chat_id
    try:
        await context.bot.send_message(
            chat_id=NOTIFY_CHAT_ID,
            text=(
                f"📣 *Убытие сотрудника*\n"
                f"👤 {full_name}\n"
                f"🏢 Отдел: {department}\n"
                f"📝 Причина: {reason}\n"
                f"🕒 {timestamp.strftime('%H:%M, %d.%m.%Y')}"
            ),
            parse_mode='Markdown'
        )
    except Exception as e:
        print(f"Ошибка при отправке уведомления: {e}")

    # Обновляем клавиатуру
    keyboard = main_button.copy()
    if user_id in ADMIN_IDS:
        keyboard += admin_button

    await update.message.reply_text(
        "✅ Информация сохранена и передана руководству Управления.\n"
        "Отмечаться при уходе домой в конце рабочего дня не нужно.",
        reply_markup=ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    )
    return ConversationHandler.END

async def report_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in ADMIN_IDS:
        await update.message.reply_text("⛔ У вас нет доступа к этой команде.")
        return
    keyboard = [
        [InlineKeyboardButton("📅 День", callback_data="day")],
        [InlineKeyboardButton("🗓 Неделя", callback_data="week")],
        [InlineKeyboardButton("📆 Месяц", callback_data="month")],
        [InlineKeyboardButton("📊 Год", callback_data="year")]
    ]
    await update.message.reply_text("Выберите период для отчёта:", reply_markup=InlineKeyboardMarkup(keyboard))


async def generate_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    period = query.data
    now = datetime.datetime.now()

    if period == "day":
        since = now.replace(hour=0, minute=1, second=0, microsecond=0)
        until = now.replace(hour=23, minute=59, second=59, microsecond=999999)
    elif period == "week":
        since = now - datetime.timedelta(days=now.weekday())
        since = since.replace(hour=0, minute=1, second=0, microsecond=0)
        until = since + datetime.timedelta(days=6, hours=23, minutes=58, seconds=59)
    elif period == "month":
        since = now.replace(day=1, hour=0, minute=1, second=0, microsecond=0)
        next_month = (since + datetime.timedelta(days=32)).replace(day=1)
        until = next_month - datetime.timedelta(seconds=1)
    elif period == "year":
        since = now.replace(month=1, day=1, hour=0, minute=1, second=0, microsecond=0)
        until = now.replace(month=12, day=31, hour=23, minute=59, second=59)
    else:
        since = now.replace(hour=0, minute=1)
        until = now

    wb = Workbook()
    ws = wb.active
    ws.title = "Отчет"
    ws.append(["№", "ФИО", "Отдел", "Причина убытия", "Дата", "Время"])

    cursor.execute("SELECT user_id, full_name, department FROM users")
    users = cursor.fetchall()

    row_num = 1
    for user in users:
        uid, name, dept = user
        cursor.execute("SELECT reason, timestamp FROM departures WHERE user_id = ? AND timestamp BETWEEN ? AND ? ORDER BY timestamp ASC", (uid, since, until))
        departs = cursor.fetchall()
        for reason, timestamp_str in departs:
            timestamp = datetime.datetime.fromisoformat(timestamp_str)
            date_str = timestamp.strftime("%d.%m.%Y")
            time_str = timestamp.strftime("%H:%M")
            ws.append([row_num, name, dept, reason, date_str, time_str])
            row_num += 1

    stream = BytesIO()
    wb.save(stream)
    stream.seek(0)
    await query.message.reply_document(document=stream, filename="report.xlsx")


async def reset(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    cursor.execute("DELETE FROM users WHERE user_id = ?", (user_id,))
    cursor.execute("DELETE FROM departures WHERE user_id = ?", (user_id,))
    conn.commit()
    await update.message.reply_text("🔁 Ваша регистрация сброшена. Чтобы пройти регистрацию заново, введите /start.", reply_markup=ReplyKeyboardMarkup([[]], resize_keyboard=True))

async def unknown(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("⚠️ Пожалуйста, используйте кнопки, предоставленные ботом. Произвольный ввод не поддерживается.")







def main():
    application = ApplicationBuilder().token("7586950102:AAHQZabnuixnT8ZkNlKkxAq3ZJCQdjnku9I").build()

    conv_handler = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            MessageHandler(filters.Regex("🚪 Сообщить об убытии с рабочего места"), handle_departure)
        ],
        states={
            REGISTER_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, register_name)],
            REGISTER_DEPARTMENT: [MessageHandler(filters.TEXT & ~filters.COMMAND, register_department)],
            WAIT_REASON: [MessageHandler(filters.TEXT & ~filters.COMMAND, save_departure)]
        },
        fallbacks=[],
        allow_reentry=True
    )

    application.add_handler(conv_handler)
    application.add_handler(CommandHandler("reset", reset))  # /reset доступен вручную
    application.add_handler(CommandHandler("report", report_command))
    application.add_handler(MessageHandler(filters.Regex("📊 Получить отчёт"), report_command))
    application.add_handler(CallbackQueryHandler(generate_report))
    application.add_handler(MessageHandler(filters.ALL, unknown))

    print("Бот запущен...")
    application.run_polling()

if __name__ == "__main__":
    main()
