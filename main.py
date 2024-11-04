import pandas as pd
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler
from io import BytesIO

# Define your bot token (replace 'YOUR_BOT_TOKEN' with the actual token)
BOT_TOKEN = '7662681407:AAHSrS0X-ksGdObc2fr1qqzn2u6xJ3nF4Hk'

# State definitions for ConversationHandler
ASK_FILE, ASK_DAYS, PROCESS_FILE = range(3)

# Start command to initiate the conversation
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Hi! Please send me the Excel file you want to process.")
    return ASK_FILE

# Step 1: Ask for the Excel file
async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['file'] = await update.message.document.get_file()  # Store the file temporarily
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
    
    # Download the file as bytes
    excel_bytes = BytesIO()
    await file.download_to_memory(excel_bytes)
    excel_bytes.seek(0)

    # Load the Excel file into a DataFrame
    data = pd.read_excel(excel_bytes)

    # ---- Add your data cleaning and transformation code here ----
    data.columns = data.iloc[12].values
    data0 = data[15:-2]
    
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

    order= data1[data1['purchase']!='overstock'][['Артикул ','Номенклатура','purchase','overstock','outofstock']]

    # Save the modified DataFrame to a new Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        order.to_excel(writer, index=False, sheet_name='Sheet1')
    output.seek(0)

    # Send the modified file back to the user
    await update.message.reply_document(
        document=output,
        filename="order_data.xlsx",
        caption="Here is your modified Excel file."
    )
    return ConversationHandler.END

# Main function to run the bot
def main() -> None:
    application = Application.builder().token(BOT_TOKEN).build()

    # Conversation handler for managing the steps
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            ASK_FILE: [MessageHandler(filters.Document.FileExtension("xlsx"), handle_file)],
            ASK_DAYS: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_days)],
            PROCESS_FILE: [MessageHandler(filters.ALL, process_file)],
        },
        fallbacks=[],
    )

    application.add_handler(conv_handler)

    # Run the bot
    application.run_polling()

if __name__ == "__main__":
    main()
