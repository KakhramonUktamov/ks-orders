import os
import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
from io import BytesIO

# Retrieve the bot token from environment variables
BOT_TOKEN = os.environ.get("TELEGRAM_TOKEN")

ASK_FILE, ASK_DAYS = range(2)  # Define the states

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Hi! Please send me the Excel file you want to process.")
    context.user_data.clear()  # Clear any previous user data
    return ASK_FILE

async def restart(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Restarting the process. Please send me the Excel file you want to process.")
    context.user_data.clear()  # Clear previous data to start fresh
    return ASK_FILE  # Directly transition to ASK_FILE state

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Process canceled. You can start again by typing /start.")
    context.user_data.clear()  # Clear any data in case they want to start again
    return ConversationHandler.END

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['file'] = await update.message.document.get_file()
    await update.message.reply_text("Now, please enter the number of days for overstock:")
    return ASK_DAYS

async def handle_days(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        # Strip leading/trailing spaces from input and handle commas or unexpected characters
        days_input = update.message.text.strip().replace(",", "")  # Remove commas if present
        days = int(days_input)  # Convert the cleaned input to an integer
        context.user_data['days'] = days
        await update.message.reply_text(f"Received! Processing your file with an overstock period of {days} days...")
        await process_file(update, context)  # Process the file
        return ConversationHandler.END  # End the conversation
    except ValueError as e:
        # Send a debug message with the exact input received for easier troubleshooting
        await update.message.reply_text(
            f"Please enter a valid number for days. (Debug: Received '{update.message.text}')"
        )
        return ASK_DAYS

async def process_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    file = context.user_data['file']
    days = context.user_data['days']
    excel_bytes = BytesIO()
    await file.download_to_memory(excel_bytes)
    excel_bytes.seek(0)
    data = pd.read_excel(excel_bytes, None, engine='openpyxl')

    # Data processing logic
    data.columns = data.iloc[12].values
    data0 = data[15:]
    data0 = data0[:-2]

    cols = ['Артикул ','Номенклатура','Дней на распродажи',
            'Остаток на конец','Средние продажи день',
           'Прошло дней от последней продажи']

    data1 = data0[cols]       
    data1 = data1.reset_index(drop=True)

    #dropping -H this kinda values:
    for_drop =[]
    for index,i in enumerate(data1['Артикул ']):
        if "-Н" in i:
            for_drop.append(index)
    data1 = data1.drop(for_drop,axis=0) 

    #dealing with numbers of selling days
    data1['Дней на распродажи'] = data1['Дней на распродажи'].str.replace(' ','')
    data1['Прошло дней от последней продажи']=data1['Прошло дней от последней продажи'].str.replace(' ','')
    data1 = data1.replace({'∞':-1})
    data1 = data1.fillna(0)
    data1[['Дней на распродажи','Прошло дней от последней продажи']] = data1[['Дней на распродажи','Прошло дней от последней продажи']].astype(int)
    data1 = data1.reset_index(drop=True)

    data1['Общый продажи период'] = data1['Средние продажи день']*days
    data1['purchase'] = 0
    data1['overstock'] = 0
    data1['outofstock'] = 0

    for i, value in enumerate(data1['Дней на распродажи']):
        if value<days and value>=0:
            data1.loc[i,'purchase'] = data1.loc[i,'Общый продажи период'] - data1.loc[i,'Остаток на конец']
        else:
            data1.loc[i,'purchase'] = 'overstock'
            data1.loc[i,'overstock'] = data1.loc[i,'Остаток на конец'] - data1.loc[i,'Общый продажи период']
            
    for i, value in enumerate(data1['Остаток на конец']):
        if value == 0:
            data1.loc[i,'outofstock'] = data1.loc[i,'Прошло дней от последней продажи']
    order= data1[['Артикул ','Номенклатура','purchase','overstock','outofstock']]


    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        order.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)

    await update.message.reply_document(document=output, filename="processed_data.xlsx", caption="Here is your processed file.")

def main() -> None:
    application = Application.builder().token(BOT_TOKEN).build()

    # Set up the conversation handler
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start), CommandHandler("restart", restart)],
        states={
            ASK_FILE: [MessageHandler(filters.Document.FileExtension("xlsx"), handle_file)],
            ASK_DAYS: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_days)],
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("restart", restart)],  # Adding fallbacks for /cancel and /restart
    )

    application.add_handler(conv_handler)
    application.add_handler(CommandHandler("restart", restart))  # Adding a standalone handler for /restart command

    # Run the bot using long polling
    application.run_polling()

if __name__ == "__main__":
    main()
