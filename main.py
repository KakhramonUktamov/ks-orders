import os
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from io import BytesIO

# Retrieve the bot token from environment variables
BOT_TOKEN = os.environ.get("TELEGRAM_TOKEN")

ASK_FILE, ASK_DAYS, ASK_BRAND, ASK_PERCENTAGE = range(4)  # Define the states

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
    # Download file and convert it to pandas DataFrame
    file = await update.message.document.get_file()
    excel_bytes = BytesIO()
    await file.download_to_memory(excel_bytes)
    excel_bytes.seek(0)
    
    try:
        # Load the data into a DataFrame and store it
        data = pd.read_excel(excel_bytes)
        context.user_data['data'] = data  # Store the DataFrame for further processing
        await update.message.reply_text("Теперь, пожалуйста, введите количество дней для overstock:")
        return ASK_DAYS
    except ValueError as e:
        await update.message.reply_text("Ошибка: Не удалось прочитать файл как допустимый файл Excel. Пожалуйста, загрузите допустимый файл .xlsx.")
        print("File read error:", e)
        return ASK_FILE  # Stay in the same state if file reading fails


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
        await query.edit_message_text("Обработка как бренд ламината. Пожалуйста, введите процент корректировки (от 0 до 1):")
        return ASK_PERCENTAGE  # Ask for the percentage if it's Laminate
    else:
        context.user_data['is_laminate'] = False
        await query.edit_message_text("Обработка без подбора характеристик.")
        # Continue processing the file without laminate adjustments
        await process_file(update, context)
        return ConversationHandler.END  # End the conversation since processing is done for non-laminate case

async def handle_percentage(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        # Get the percentage from user input
        percentage = float(update.message.text.strip())

        # Check if the percentage is between 0 and 1
        if 0 <= percentage <= 1:
            context.user_data['percentage'] = percentage  # Store percentage for later use
            await update.message.reply_text(f"Процент принят: {percentage}. Продолжаем обработку с обновленными данными.")
            
            # Proceed with further processing
            await process_file(update, context)
            return ConversationHandler.END
        else:
            await update.message.reply_text("Пожалуйста, введите корректный процент от 0 до 1.")
            return ASK_PERCENTAGE  # Stay in the same state until valid input
    except ValueError:
        await update.message.reply_text("Пожалуйста, введите допустимое число для процента.")
        return ASK_PERCENTAGE



async def process_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    # Identify if we have an update from a callback query or a regular message
    message = update.message if update.message else update.callback_query.message
    
    try:
       # Access the stored file and DataFrame from user_data
        data = context.user_data.get('data')  # This is the pandas DataFrame already stored
        days = context.user_data.get('days')
        is_laminate = context.user_data.get('is_laminate', False)
        percentage = context.user_data.get('percentage', 1)
        
        if data is None:
            await message.reply_text("Ошибка: Данные не были загружены.")
            return ConversationHandler.END  # Exit if data is not available
    
        
        
            # Adjust daily sales based on the percentage
            
        # Data processing logic
        data0 = data[2:]
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

        data1['Общый продажи период'] = data1['Средние продажи день'] * days
        data1['helper'] = 0.0  # Set as float to allow numerical and 'overstock' entries
        data1['overstock'] = 0.0
        data1['outofstock'] = 0
        data1['В пути'] = 0
        
        # Ensure numerical columns are float-compatible
        if is_laminate:
            # Adjust the average daily sales if it's Laminate
            data1['Общый продажи период'] = data1['Средние продажи день'] * days * percentage
            #making a class column 
        
        

        for i, value in enumerate(data1['Дней на распродажи']):
            if value <= days and value >= 0:
                # Calculate purchase as float
                data1.loc[i, 'helper'] = float(data1.loc[i, 'Общый продажи период'] - data1.loc[i, 'Остаток на конец'])
            else:
                # Set 'overstock' string, handled by float-compatible column
                data1.loc[i, 'helper'] = 0
                data1.loc[i, 'overstock'] = float(data1.loc[i, 'Остаток на конец'] - data1.loc[i, 'Общый продажи период'])

        for i, value in enumerate(data1['Остаток на конец']):
            if value <= 50:
                data1.loc[i, 'outofstock'] = data1.loc[i, 'Прошло дней от последней продажи'] * 
                    data1.loc[i, 'Средние продажи день']*percentage - data1.loc[i, 'Остаток на конец']

        data1['Рекомендательный Заказ'] = 'helper - on_the_way'

        features = ['ЕMR','EMR','YEL','WHT','ULT','SF','RUB','RED','PG','ORN','NC',
                    'LM','LAG','IND','GRN','GREY','FP STNX','FP PLC','FP NTR','CHR',
                    'BLU','BLA','AMB']

        def find_feature(text):
            for feature in features:
                if pd.notna(text) and feature in text:
                    return feature
            return "No Match"

        # Apply the function to the column containing text (e.g., 'Description')
        data1['Коллекция'] = data1['Номенклатура'].apply(find_feature)
        # Creating the `purchase_df` with the necessary columns
        purchase_df = data1[['Артикул ', 'Номенклатура', 'Collection', 'helper','В пути','Рекомендательный Заказ']]
        # Add the `on_the_way` column with a default value of 0
        
        #Separate DataFrames for each sheet]
        overstock_df = data1[['Артикул ', 'Номенклатура', 'Коллекция', 'overstock']]
        outofstock_df = data1[['Артикул ', 'Номенклатура', 'outofstock']]

        # Add USD calculation column to outofstock_df
        outofstock_df['USD of outofstock'] = ''  # Multiplied by E1 value placeholder

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write each DataFrame to its respective sheet
            purchase_df.to_excel(writer, sheet_name='Рекомендательный Заказ', index=False)
            overstock_df.to_excel(writer, sheet_name='Overstock', index=False)
            outofstock_df.to_excel(writer, sheet_name='OutOfStock', index=False)

            # Access the OutOfStock sheet to add fixed E1 value
            workbook = writer.book
            worksheet = writer.sheets['OutOfStock']
            worksheet.write('E1', 1)  # Fixed cell value for USD multiplier
            worksheet1 = writer.sheets["Рекомендательный Заказ"]

            # Write the formula for each row in the 'USD of outofstock' column
            for row_num in range(1, len(outofstock_df) + 1):
                formula = f'=C{row_num + 1}*$E$1'  # C column for outofstock, $E$1 for fixed value
                worksheet.write_formula(row_num, 3, formula)  # Write formula in D column (USD of outofstock)
            
            # Write formulas to `total_purchase` column for dynamic calculation
            for row_num in range(1, len(purchase_df) + 1):  # Starting from row 1 to avoid headers
                worksheet1.write_formula(
                    row_num, 5, f'=MAX(D{row_num + 1} - E{row_num + 1}, 0)'
                )  # Ensure the result is not less than 0

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
            ASK_PERCENTAGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_percentage)], 
        },
        fallbacks=[CommandHandler("cancel", cancel), CommandHandler("restart", restart)],  # Adding fallbacks for /cancel and /restart
    )

    application.add_handler(conv_handler)
    application.add_handler(CommandHandler("restart", restart))  # Adding a standalone handler for /restart command

    # Run the bot using long polling
    application.run_polling()

if __name__ == "__main__":
    main()
