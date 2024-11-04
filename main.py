import os
import pandas as pd
from flask import Flask, request
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
from io import BytesIO

# Initialize Flask app
app = Flask(__name__)

# Define environment variables and constants
PORT = int(os.environ.get("PORT", "8443"))
BOT_TOKEN = os.environ.get("TELEGRAM_TOKEN")  # Store your bot token in Heroku Config Vars
APP_URL = os.environ.get("APP_URL")           # Store your Heroku app URL in Config Vars

# State definitions for ConversationHandler
ASK_FILE, ASK_DAYS, PROCESS_FILE = range(3)

# Start command to initiate the conversation
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Hi! Please send me the Excel file you want to process.")
    return ASK_FILE

# Step 1: Ask for the Excel file
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['file'] = await update.message.document.get_file()
    await update.message.reply_text("Now, please enter the number of days for overstock:")
    return ASK_DAYS

# Step 2: Handle the days input
async def handle_days(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        days = int(update.message.text)
        context.user_data['days'] = days
        await update.message.reply_text(f"Received! Processing your file with an overstock period of {days} days...")
        return await process_file(update, context)
    except ValueError:
        await update.message.reply_text("Please enter a valid number for days.")
        return ASK_DAYS

# Step 3: Process the Excel file with the days value
async def process_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    file = context.user_data['file']
    days = context.user_data['days']
    excel_bytes = BytesIO()
    await file.download_to_memory(excel_bytes)
    excel_bytes.seek(0)
    data = pd.read_excel(excel_bytes)

    # Data processing
    data.columns = data.iloc[12].values
    data0 = data[15:-2]
    cols = ['Артикул ','Номенклатура','Дней на распродажи', 'Остаток на конец', 'Средние продажи день', 'Прошло дней от последней продажи']
    data1 = data0[cols].reset_index(drop=True)

    # Dropping '-Н' values and cleaning data
    data1 = data1[~data1['Артикул '].str.contains("-Н")]
    data1[['Дней на распродажи', 'Прошло дней от последней продажи']] = data1[['Дней на распродажи', 'Прошло дней от последней продажи']].apply(lambda x: x.str.replace(' ', '')).replace({'∞': -1}).fillna(0).astype(int)
    data1['Общый продажи период'] = data1['Средние продажи день'] * days
    data1['purchase'] = data1.apply(lambda row: row['Общый продажи период'] - row['Остаток на конец'] if row['Дней на распродажи'] < days else 'overstock', axis=1)
    data1['overstock'] = data1.apply(lambda row: row['Остаток на конец'] - row['Общый продажи период'] if row['purchase'] == 'overstock' else 0, axis=1)
    data1['outofstock'] = data1.apply(lambda row: row['Прошло дней от последней продажи'] if row['Остаток на конец'] == 0 else 0, axis=1)
    order = data1[data1['purchase'] != 'overstock'][['Артикул ','Номенклатура','purchase','overstock','outofstock']]

    # Save the modified DataFrame to a new Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        order.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)

    # Send the modified file back to the user
    await update.message.reply_document(document=output, filename="order_data.xlsx", caption="Here is your modified Excel file.")
    return ConversationHandler.END

# Configure and run the bot using webhooks
application = Application.builder().token(BOT_TOKEN).build()
application.add_handler(ConversationHandler(
    entry_points=[CommandHandler("start", start)],
    states={
        ASK_FILE: [MessageHandler(filters.Document.FileExtension("xlsx"), handle_file)],
        ASK_DAYS: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_days)],
        PROCESS_FILE: [MessageHandler(filters.ALL, process_file)],
    },
    fallbacks=[],
))

# Webhook route
@app.route('/webhook', methods=['POST'])
async def webhook():
    update = Update.de_json(request.get_json(), application.bot)
    await application.process_update(update)
    return 'OK'

# Function to set webhook
async def set_webhook():
    await application.bot.set_webhook(f"{APP_URL}/webhook")

# Start the Flask app with webhook
if __name__ == '__main__':
    import asyncio
    loop = asyncio.get_event_loop()
    loop.run_until_complete(set_webhook())
    app.run(host='0.0.0.0', port=PORT)
