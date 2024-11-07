import os
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from io import BytesIO

# Retrieve the bot token from environment variables
BOT_TOKEN = os.environ.get("TELEGRAM_TOKEN")

ASK_FILE, ASK_DAYS, ASK_BRAND = range(3)  # Define the states

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Ассалому Алайкум! Пожалуйста, отправьте мне Excel файл, который вы хотите обработать.")
    context.user_data.clear()  # Clear any previous user data
    return ASK_FILE

async def restart(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Перезапуск процесса. Пожалуйста, отправьте мне Excel файл, который вы хотите обработать.")
    context.user_data.clear()  # Clear previous data to start fresh
    return ASK_FILE  # Directly transition to ASK_FILE state

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Процесс отменен. Вы можете начать заново, набрав /start.")
    context.user_data.clear()  # Clear any data in case they want to start again
    return ConversationHandler.END

async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data['file'] = await update.message.document.get_file()
    await update.message.reply_text("Теперь, пожалуйста, введите количество дней для overstock:")
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
        
    except ValueError as e:
        # Send a debug message with the exact input received for easier troubleshooting
        await update.message.reply_text(
            f"Пожалуйста, введите допустимое число для дней. (Отладка: получено '{update.message.text}')"
        )
        return ASK_DAYS

async def handle_brand(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    
    # Check user's choice and set a flag in user_data
    if query.data == "yes":
        context.user_data['is_laminate'] = True
        await query.edit_message_text("Обработка как бренд ламината с подбором характеристик.")
    else:
        context.user_data['is_laminate'] = False
        await query.edit_message_text("Обработка без подбора характеристик.")
    
    # Proceed to file processing
    await process_file(update, context)
    return ConversationHandler.END

async def process_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    # Identify if we have an update from a callback query or a regular message
    message = update.message if update.message else update.callback_query.message
    
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
            await message.reply_text("Ошибка: Не удалось прочитать файл как допустимый файл Excel. Пожалуйста, загрузите допустимый файл .xlsx.")
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
            if value <= days and value >= 0:
                # Calculate purchase as float
                data1.loc[i, 'purchase'] = float(data1.loc[i, 'Общый продажи период'] - data1.loc[i, 'Остаток на конец'])
            else:
                # Set 'overstock' string, handled by float-compatible column
                data1.loc[i, 'purchase'] = 'overstock'
                data1.loc[i, 'overstock'] = float(data1.loc[i, 'Остаток на конец'] - data1.loc[i, 'Общый продажи период'])

        for i, value in enumerate(data1['Остаток на конец']):
            if value == 0:
                data1.loc[i, 'outofstock'] = data1.loc[i, 'Прошло дней от последней продажи'] * data1.loc[i, 'Средние продажи день']

        #making a class column 
        features = ['ЕMR','EMR','YEL','WHT','ULT','SF','RUB','RED','PG','ORN','NC',
                    'LM','LAG','IND','GRN','GREY','FP STNX','FP PLC','FP NTR','CHR',
                    'BLU','BLA','AMB']

        def find_feature(text):
            for feature in features:
                if pd.notna(text) and feature in text:
                    return feature
            return "No Match"

        # Apply the function to the column containing text (e.g., 'Description')
        data1['Collection'] = data1['Номенклатура'].apply(find_feature)

        # Separate DataFrames for each sheet
        purchase_df = data1[['Артикул ', 'Номенклатура', 'Collection', 'purchase']]
        overstock_df = data1[['Артикул ', 'Номенклатура', 'Collection', 'overstock']]
        outofstock_df = data1[['Артикул ', 'Номенклатура', 'outofstock']]

        # Add USD calculation column to outofstock_df
        outofstock_df['USD of outofstock'] = ''  # Multiplied by E1 value placeholder

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write each DataFrame to its respective sheet
            purchase_df.to_excel(writer, sheet_name='Purchase', index=False)
            overstock_df.to_excel(writer, sheet_name='Overstock', index=False)
            outofstock_df.to_excel(writer, sheet_name='OutOfStock', index=False)

            # Access the OutOfStock sheet to add fixed E1 value
            workbook = writer.book
            worksheet = writer.sheets['OutOfStock']
            worksheet.write('E1', 1)  # Fixed cell value for USD multiplier

            # Write the formula for each row in the 'USD of outofstock' column
            for row_num in range(1, len(outofstock_df) + 1):
                formula = f'=C{row_num + 1}*$E$1'  # C column for outofstock, $E$1 for fixed value
                worksheet.write_formula(row_num, 3, formula)  # Write formula in D column (USD of outofstock)

        output.seek(0)

        await message.reply_document(document=output, filename="processed_data.xlsx", caption="Вот ваш обработанный файл.")

    except Exception as e:
        await message.reply_text("Произошла непредвиденная ошибка при обработке файла.")
        print("Processing error:",e)


def main() -> None:
    application = Application.builder().token(BOT_TOKEN).build()

    # Set up the conversation handler
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start), CommandHandler("restart", restart)],
        states={
            ASK_FILE: [MessageHandler(filters.Document.FileExtension("xlsx"), handle_file)],
            ASK_DAYS: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_days)],
            ASK_BRAND: [CallbackQueryHandler(handle_brand)],
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("restart", restart)],  # Adding fallbacks for /cancel and /restart
    )

    application.add_handler(conv_handler)
    application.add_handler(CommandHandler("restart", restart))  # Adding a standalone handler for /restart command

    # Run the bot using long polling
    application.run_polling()

if __name__ == "__main__":
    main()
