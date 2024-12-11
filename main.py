import os
import logging
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from io import BytesIO

# Configure logging
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO,
    filename="bot_activity.log"  # Logs will be saved to this file
)

logger = logging.getLogger(__name__)


# Retrieve the bot token from environment variables
BOT_TOKEN = os.environ.get("TELEGRAM_TOKEN")

ALLOWED_NUMBERS = ["+998916919534", "+998958330373",
                   "+998884758000","+998900212141","+998998449669"]  # Replace with your company's authorized phone numbers

ASK_FILE, ASK_DAYS, ASK_BRAND, ASK_PERCENTAGE = range(4)  # Define the states



def normalize_phone_number(phone_number: str) -> str:
    """Normalize phone numbers to international format."""
    phone_number = "".join(c for c in phone_number if c.isdigit() or c == "+")
    # Ensure all numbers start with '+'
    if not phone_number.startswith("+"):
        phone_number = "+" + phone_number
    return phone_number

async def handle_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle phone numbers sent via the 'Share Phone Number' button."""
    if update.message.contact:  # Phone number shared via "Share Phone Number" button
        phone_number = normalize_phone_number(update.message.contact.phone_number)

        # Log the received phone number
        logger.info(f"Phone number received via button: {phone_number}")
      
        # Check if the phone number is in the allowed list
        if phone_number in ALLOWED_NUMBERS:
            context.user_data['verified'] = True  # Mark the user as verified
            await update.message.reply_text(
                "Доступ предоставлен! ✅. Добро пожаловать в бот от KS Group! Вы подтверждены.\n"
                "Пожалуйста, отправьте Excel файл, который вы хотите обработать."
            )
            return ASK_FILE  # Proceed to file processing
        else:
            logger.warning(f"Unauthorized access attempt with phone number: {phone_number}")
            await update.message.reply_text(
                "Доступ запрещен! ❌. Ваш номер телефона не авторизован для использования этого бота."
            )
            return ConversationHandler.END
    else:              # User typed their phone number manually
        logger.warning(f"Manual input detected: {update.message.text.strip()}")
        await update.message.reply_text(
            "Пожалуйста, используйте кнопку 'Поделиться номером телефона', чтобы отправить свой номер для подтверждения."
        )
        return ASK_FILE  # Stay in the current state, waiting for correct input


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Start the bot and request phone verification if needed."""
    # Check if the user has already verified their phone number
    if 'verified' in context.user_data and context.user_data['verified']:
        await update.message.reply_text("Пожалуйста, отправьте мне Excel файл, который вы хотите обработать.")
        return ASK_FILE

    # Prompt for phone number verification
    keyboard = [[KeyboardButton("Share Phone Number", request_contact=True)]]
    reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
    
    await update.message.reply_text(
        "Для обеспечения безопасности, пожалуйста, поделитесь своим номером телефона для подтверждения вашей личности.",
        reply_markup=reply_markup
    )
    logger.info(f"Start command received from user: {update.message.chat.username}")
    return ASK_FILE  # Move to the file handling state once verified 


async def restart(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Перезапуск процесса. Пожалуйста, отправьте мне Excel файл, который вы хотите обработать.")
    context.user_data.clear()  # Clear previous data to start fresh
    return ASK_FILE  # Directly transition to ASK_FILE state

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    logger.info(f"Cancel command received from user: {update.message.chat.username}")
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

    except Exception as e:
        await update.message.reply_text("Произошла ошибка при сохранении файла.")
        print("File save error:", e)
        return ASK_FILE


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
        data1['В Пути'] = 0
        
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
                data1.loc[i, 'outofstock'] = data1.loc[i, 'Прошло дней от последней продажи'] * data1.loc[i, 'Средние продажи день']*percentage - data1.loc[i, 'Остаток на конец']

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
        purchase_df = data1[['Артикул ', 'Номенклатура', 'Коллекция', 'helper','В Пути','Рекомендательный Заказ']]
        # Add the `on_the_way` column with a default value of 0
        
        #Separate DataFrames for each sheet]
        overstock_df = data1[['Артикул ', 'Номенклатура', 'Коллекция', 'overstock']]
        outofstock_df = data1[['Артикул ', 'Номенклатура', 'outofstock']]

        # Add USD calculation column to outofstock_df
        outofstock_df['USD of outofstock'] = ''  # Multiplied by E1 value placeholder

        on_the_way = pd.DataFrame(columns = ['Артикул ', 'Номенклатура','В Пути'])

        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            # Write each DataFrame to its respective sheet
            purchase_df.to_excel(writer, sheet_name='Рекомендательный Заказ', index=False)
            overstock_df.to_excel(writer, sheet_name='Overstock', index=False)
            outofstock_df.to_excel(writer, sheet_name='OutOfStock', index=False)
            on_the_way.to_excel(writer, sheet_name='В Пути', index=False)

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

            for row_num in range(1, len(purchase_df) + 1):
                worksheet1.write_formula(row_num - 1, 4, f'=iferror(VLOOKUP(A{row_num}, \'В Пути\'!A:C, 3, FALSE),0)')


        output.seek(0)

        await message.reply_document(document=output, filename="processed_data.xlsx", caption="📎Вот ваш обработанный файл.")

    except Exception as e:
        await message.reply_text("Произошла непредвиденная ошибка при обработке файла.")
        print("Processing error:",e)


def main() -> None:
    application = Application.builder().token(BOT_TOKEN).build()

    # Set up the conversation handler
    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            ASK_FILE: [
                MessageHandler(filters.Document.FileExtension("xlsx"), handle_file),
                MessageHandler(filters.CONTACT, handle_phone),
                MessageHandler(filters.TEXT & ~filters.COMMAND, handle_phone)
            ],
            ASK_DAYS: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_days)],
            ASK_BRAND: [CallbackQueryHandler(handle_brand)],
            ASK_PERCENTAGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_percentage)], 
        },
        fallbacks=[CommandHandler("cancel", cancel)],  # Adding fallbacks for /cancel and /restart
    )

    application.add_handler(conv_handler)

    # Run the bot using long polling
    application.run_polling()

if __name__ == "__main__":
    main()
