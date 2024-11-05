import os
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
from io import BytesIO
import requests

BOT_TOKEN = os.environ.get("TELEGRAM_TOKEN")
#WEBHOOK_URL = f'https://ks-orders.onrender.com/{BOT_TOKEN}'  # Change this to your Render app URL


ASK_FILE, ASK_DAYS = range(2)  # We only need these two states

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Hi! Please send me the Excel file you want to process.")
    context.user_data.clear()  # Resetting user data
    return ASK_FILE

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['file'] = await update.message.document.get_file()
    await update.message.reply_text("Now, please enter the number of days for overstock:")
    return ASK_DAYS

async def handle_days(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        days = int(update.message.text)
        context.user_data['days'] = days
        await update.message.reply_text(f"Received! Processing your file with an overstock period of {days} days...")
        await process_file(update, context)  # Call process_file directly
        return ConversationHandler.END
    except ValueError:
        await update.message.reply_text("Please enter a valid number for days.")
        return ASK_DAYS

async def process_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    file = context.user_data['file']
    days = context.user_data['days']
    excel_bytes = BytesIO()
    await file.download_to_memory(excel_bytes)
    excel_bytes.seek(0)
    data = pd.read_excel(excel_bytes)

    # Add your data processing code here
    data.columns = data.iloc[12].values
    data0 = data[15:-2]
    cols = ['Артикул ', 'Номенклатура', 'Дней на распродажи', 'Остаток на конец', 'Средние продажи день', 'Прошло дней от последней продажи']
    data1 = data0[cols].reset_index(drop=True)

    # Dropping '-Н' values and cleaning data
    data1 = data1[~data1['Артикул '].str.contains("-Н")]
    data1[['Дней на распродажи', 'Прошло дней от последней продажи']] = data1[['Дней на распродажи', 'Прошло дней от последней продажи']].apply(lambda x: x.str.replace(' ', '')).replace({'∞': -1}).fillna(0).astype(int)
    data1['Общый продажи период'] = data1['Средние продажи день'] * days
    data1['purchase'] = data1.apply(lambda row: row['Общый продажи период'] - row['Остаток на конец'] if row['Дней на распродажи'] < days else 'overstock', axis=1)
    data1['overstock'] = data1.apply(lambda row: row['Остаток на конец'] - row['Общый продажи период'] if row['purchase'] == 'overstock' else 0, axis=1)
    data1['outofstock'] = data1.apply(lambda row: row['Прошло дней от последней продажи'] if row['Остаток на конец'] == 0 else 0, axis=1)
    order = data1[['Артикул ', 'Номенклатура', 'purchase', 'overstock', 'outofstock']]

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        order.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)

    await update.message.reply_document(document=output, filename="processed_data.xlsx", caption="Here is your processed file.")



def main() -> None:
    application = Application.builder().token(BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            ASK_FILE: [MessageHandler(filters.Document.FileExtension("xlsx"), handle_file)],
            ASK_DAYS: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_days)],
        },
        fallbacks=[],
    )

    application.add_handler(conv_handler)

    # Run the bot
    application.run_polling()

if __name__ == "__main__":
    main()
