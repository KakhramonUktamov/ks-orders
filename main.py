import os
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from io import BytesIO

# Retrieve the bot token from environment variables
BOT_TOKEN = os.environ.get("TELEGRAM_TOKEN")

ASK_FILE, ASK_DAYS, ASK_BRAND = range(3)  # Define the states

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
        days_input = update.message.text.strip() # Remove commas if present
        days = int(days_input)  # Convert the cleaned input to an integer
        context.user_data['days'] = days

        keyboard = [
            [InlineKeyboardButton("Да", callback_data="yes"), InlineKeyboardButton("Нет", callback_data="no")]
        ]
        reply_markup = InlineKeyboardMarkup(keyboard)
        await update.message.reply_text("Этот бренд Ламинат?", reply_markup=reply_markup)
        return ASK_BRAND
        
        await update.message.reply_text(f"Received! Processing your file with an overstock period of {days} days...")
        await process_file(update, context)  # Process the file
        return ConversationHandler.END  # End the conversation
    except ValueError as e:
        # Send a debug message with the exact input received for easier troubleshooting
        await update.message.reply_text(
            f"Please enter a valid number for days. (Debug: Received '{update.message.text}')"
        )
        return ASK_DAYS

async def handle_brand(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    
    # Check user's choice and set a flag in user_data
    if query.data == "yes":
        context.user_data['is_laminate'] = True
        await query.edit_message_text("Processing as Laminate brand with feature matching.")
    else:
        context.user_data['is_laminate'] = False
        await query.edit_message_text("Processing without feature matching.")
    
    # Proceed to file processing
    await process_file(update, context)
    return ConversationHandler.END

async def process_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    try:
        file = context.user_data['file']
        days = context.user_data['days']
        is_laminate = context.user_data.get('is_laminate', False)
        
        excel_bytes = BytesIO()
        await file.download_to_memory(excel_bytes)
        excel_bytes.seek(0)
        
        try:
            data = pd.read_excel(excel_bytes)
        except ValueError as e:
            await update.message.reply_text("Error: Unable to read the file as a valid Excel file. Please upload a valid .xlsx file.")
            print("File read error:", e)
            return  # Exit the function if the file is not valid
    

        # Data processing logic
        data.columns = data.iloc[12].values
        data0 = data[15:]
        data0 = data0[:-2]

        cols = ['Артикул ', 'Номенклатура', 'Дней на распродажи',
                'Остаток на конец', 'Средние продажи день',
                'Прошло дней от последней продажи']

        data1 = data0[cols]
        data1 = data1.reset_index(drop=True)

        # Dropping '-H' values
        data1 = data1.drop(data1[data1['Артикул '].astype(str).str.contains("-Н", na=False)].index)
        
        # Dealing with numbers of selling days
        data1['Дней на распродажи'] = data1['Дней на распродажи'].str.replace(' ', '')
        data1['Прошло дней от последней продажи'] = data1['Прошло дней от последней продажи'].str.replace(' ', '')
        data1 = data1.replace({'∞': -1})
        data1 = data1.fillna(0)
        data1[['Дней на распродажи', 'Прошло дней от последней продажи']] = data1[['Дней на распродажи', 'Прошло дней от последней продажи']].astype(int)
        data1 = data1.reset_index(drop=True)

        # Ensure numerical columns are float-compatible
        data1['Общый продажи период'] = data1['Средние продажи день'] * days
        data1['purchase'] = 0.0  # Set as float to allow numerical and 'overstock' entries
        data1['overstock'] = 0.0
        data1['outofstock'] = 0

        for i, value in enumerate(data1['Дней на распродажи']):
            if value < days and value >= 0:
                # Calculate purchase as float
                data1.loc[i, 'purchase'] = float(data1.loc[i, 'Общый продажи период'] - data1.loc[i, 'Остаток на конец'])
            else:
                # Set 'overstock' string, handled by float-compatible column
                data1.loc[i, 'purchase'] = 'overstock'
                data1.loc[i, 'overstock'] = float(data1.loc[i, 'Остаток на конец'] - data1.loc[i, 'Общый продажи период'])

        for i, value in enumerate(data1['Остаток на конец']):
            if value == 0:
                data1.loc[i, 'outofstock'] = data1.loc[i, 'Прошло дней от последней продажи'] * data1.iloc[i, 'Средние продажи день']

        #making a class column 
        def find_feature(text):
            for feature in features:
                if pd.notna(text) and feature in text:
                    return feature
            return "No Match"

        features = ['EMR','YEL','WHT','ULT','SF','RUB','RED','PG','ORN','NC',
                    'LM','LAG','IND','GRN','GREY','FP STNX','FP PLC','FP NTR','CHR',
                    'BLU','BLA','AMB']
        # Apply the function to the column containing text (e.g., 'Description')
        data1['Class'] = data['Номенклатура'].apply(find_feature)

                                                            
        #Select final columns for output
        order = data1[['Артикул ', 'Номенклатура','Class', 'purchase', 'overstock', 'outofstock']]


        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            order.to_excel(writer, index=False, sheet_name='Sheet1')
        output.seek(0)

        await update.message.reply_document(document=output, filename="processed_data.xlsx", caption="Here is your processed file.")

    except Exception as e:
        await update.message.reply_text("An unexpected error occurred while processing the file.")
        print("Processing error:",e)


def main() -> None:
    application = Application.builder().token(BOT_TOKEN).build()

    # Set up the conversation handler
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start), CommandHandler("restart", restart)],
        states={
            ASK_FILE: [MessageHandler(filters.Document.FileExtension("xlsx"), handle_file)],
            ASK_DAYS: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_days)],
            ASK_BRAND: [MessageHandler(CallbackQueryHandler(handle_brand))],
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("restart", restart)],  # Adding fallbacks for /cancel and /restart
    )

    application.add_handler(conv_handler)
    application.add_handler(CommandHandler("restart", restart))  # Adding a standalone handler for /restart command

    # Run the bot using long polling
    application.run_polling()

if __name__ == "__main__":
    main()
