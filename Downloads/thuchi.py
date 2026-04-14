from telegram.ext import Application, CommandHandler, MessageHandler, filters
from openpyxl import load_workbook
from datetime import datetime

FILE_PATH = "thu_chi.xlsx"

def update_excel(file_path, user_input):
    wb = load_workbook(file_path)
    ws = wb.active

    rows = user_input.split(".")
    for row in rows:
        parts = row.strip().split(",")
        if len(parts) < 3:
            continue

        today = datetime.now().strftime("%d/%m/%Y")
        thu = int(parts[0]) if parts[0].strip() else 0
        chi = int(parts[1]) if parts[1].strip() else 0
        noidung = parts[2].strip()

        ws.append([today, thu, chi, noidung])

    wb.save(file_path)

async def start(update, context):
    await update.message.reply_text(
        "Xin chào! Hãy gửi dữ liệu theo dạng:\n"
        "93000,20000, thu-chi. 61000,10000, ăn-uống. 100000,5000, làm-chơi\n\n"
        "Mỗi mục cách nhau bằng dấu chấm (.)"
    )

async def add_data(update, context):
    text = update.message.text.strip()
    update_excel(FILE_PATH, text)
    await update.message.reply_text("✅ Đã ghi dữ liệu vào file Excel!")

def main():
    app = Application.builder().token("8658634059:AAHdoI8vt7E_yR8tevuFE9vzgj4Wh2bvkkM").build()

    # Lệnh /start để gửi hướng dẫn
    app.add_handler(CommandHandler("start", start))

    # Nhận mọi tin nhắn text để thêm dữ liệu
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, add_data))

    print("Bot đang chạy...")
    app.run_polling()

if __name__ == "__main__":
    main()
