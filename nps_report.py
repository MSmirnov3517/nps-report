import imaplib
import email
import os
from datetime import datetime
import pandas as pd
import re
from telegram import Bot
import asyncio
import random

# Настройки
EMAIL = "MsL-2025@mail.ru"
IMAP_SERVER = "imap.mail.ru"
PASSWORD = os.getenv("EMAIL_PASSWORD")
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")

# Список эмодзи
EMOJIS = ["✨", "🌟", "🚀", "💡", "🎉", "🔥", "🌈", "⚡", "🍀", "🌼"]

# Подключение к почте
def connect_to_email():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, PASSWORD)
    mail.select("INBOX")
    return mail

# Поиск письма с NPS за сегодня
def fetch_email(mail):
    today = datetime.now().strftime("%d-%b-%Y")
    _, message_numbers = mail.search(None, f'(SINCE "{today}" SUBJECT "NPS")')
    for num in message_numbers[0].split():
        _, msg_data = mail.fetch(num, "(RFC822)")
        email_body = msg_data[0][1]
        email_msg = email.message_from_bytes(email_body)
        for part in email_msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get("Content-Disposition") is None:
                continue
            filename = part.get_filename()
            if "NPS" in filename and filename.endswith(".xlsx"):
                filepath = os.path.join("temp", filename)
                os.makedirs("temp", exist_ok=True)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                return filepath, filename
    return None, None

# Обработка Excel и расчет NPS
def process_excel(filepath, filename):
    # Извлечение даты из названия файла (например, "NPS 02.06.2025.xlsx")
    date_str = re.search(r"\d{2}\.\d{2}\.\d{4}", filename).group()
    report_date = datetime.strptime(date_str, "%d.%m.%Y").strftime("%Y-%m-%d")  # Дата "вчера"
    month_start = datetime.strptime(date_str, "%d.%m.%Y").replace(day=1).strftime("%Y-%m-%d")  # Начало месяца

    df = pd.read_excel(filepath)
    # Фильтр за вчера и с начала месяца
    df["Дата"] = pd.to_datetime(df["Дата"])
    df_yesterday = df[df["Дата"] == report_date]
    df_month = df[df["Дата"] >= month_start]

    # Подсчет оценок
    ratings_yesterday = df_yesterday["Оценка"].value_counts().to_dict()
    ratings_month = df_month["Оценка"].value_counts().to_dict()

    # NPS: (5 - (3+2+1)) / (5+4+3+2+1) * 100
    def calculate_nps(ratings):
        r5 = ratings.get(5, 0)
        r4 = ratings.get(4, 0)
        r3 = ratings.get(3, 0)
        r2 = ratings.get(2, 0)
        r1 = ratings.get(1, 0)
        total = r5 + r4 + r3 + r2 + r1
        if total == 0:
            return 0
        return round(((r5 - (r3 + r2 + r1)) / total) * 100)

    nps_yesterday = calculate_nps(ratings_yesterday)
    nps_month = calculate_nps(ratings_month)

    # Благодарности (все комментарии с оценкой 5 за день)
    thanks = df_yesterday[df_yesterday["Оценка"] == 5]["Комментарий"].dropna()
    thanks_text = "\n".join(thanks) if not thanks.empty else ""

    # Жалобы (оценки 1, 2, 3)
    complaints = df_yesterday[df_yesterday["Оценка"].isin([1, 2, 3])]
    complaints_text = ""
    for _, row in complaints.iterrows():
        complaints_text += f"{row['Категория']}  {row['Оценка']}  {row['Id жалобы']}  {row['Комментарий']}\n"

    return {
        "nps_month": nps_month,
        "nps_yesterday": nps_yesterday,
        "count_5": ratings_yesterday.get(5, 0),
        "thanks": thanks_text,
        "complaints": complaints_text or "Нет жалоб"
    }

# Отправка отчета в Telegram
async def send_report(data=None):
    bot = Bot(token=TELEGRAM_TOKEN)
    if data:
        report = (
            f"NPS с начала месяца - {data['nps_month']}% {random.choice(EMOJIS)}\n"
            f"\n"  # Пустая строка
            f"NPS за вчера - {data['nps_yesterday']}% {random.choice(EMOJIS)}\n"
            f"Количество 5 - {data['count_5']} {random.choice(EMOJIS)}\n"
        )
        if data["thanks"]:
            report += f"Благодарности гостей - {data['thanks']} 👍 {random.choice(EMOJIS)}\n"
        report += f"Жалобы\n{data['complaints']} {random.choice(EMOJIS)}"
    else:
        report = f"Отчет за вчера не пришел или не содержит нужного файла 📭"
    await bot.send_message(chat_id=CHAT_ID, text=report)

# Основной процесс
async def main():
    mail = connect_to_email()
    filepath, filename = fetch_email(mail)
    if filepath and filename:
        data = process_excel(filepath, filename)
        await send_report(data)
        os.remove(filepath)  # Удаляем временный файл
    else:
        await send_report()  # Уведомление об отсутствии отчета
    mail.logout()

if __name__ == "__main__":
    asyncio.run(main())
