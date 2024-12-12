import os
from datetime import datetime
import logging
import pandas as pd
import json
from collections import defaultdict
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from io import BytesIO

import logging

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

ASK_FILE, ASK_DAYS, ASK_BRAND, ASK_PERCENTAGE, ASK_PHONE, MAIN_MENU, PRODUCT_ORDER, ADMIN_PANEL, HELP_MENU = range(9)  # Define the states

# Dictionary to track user activity
user_activity = defaultdict(lambda: {
    "usage_count": 0, 
    "phone_number": None,
    "last_used": None
})
# Admin and allowed phone numbers
ADMIN_PHONE = "+998916919534"
ALLOWED_NUMBERS = ["+998916919534"]

USER_ACTIVITY_FILE = "user_activity.json"
BOT_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ADMIN_TELEGRAM_ID = os.environ.get("ADMIN_TELEGRAM_ID")

# Load previous activity from file
if os.path.exists(USER_ACTIVITY_FILE):
    with open(USER_ACTIVITY_FILE, "r") as file:
        user_activity.update(json.load(file))



def normalize_phone_number(phone_number: str) -> str:
    """Normalize phone numbers to international format."""
    phone_number = "".join(c for c in phone_number if c.isdigit() or c == "+")
    # Ensure all numbers start with '+'
    if not phone_number.startswith("+"):
        phone_number = "+" + phone_number
    return phone_number


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user = update.message.from_user
    context.user_data.clear()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    user_activity[user.username]["usage_count"] += 1
    user_activity[user.username]["last_used"] = now
    user_activity[user.username]["phone_number"] = None
    with open(USER_ACTIVITY_FILE, "w") as file:
        json.dump(user_activity, file)

    keyboard = [[KeyboardButton("Share Phone Number", request_contact=True)]]
    reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
    await update.message.reply_text("Please share your phone number for verification.", reply_markup=reply_markup)
    return ASK_PHONE


async def handle_phone(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    user = update.message.from_user
    if update.message.contact:
        phone_number = normalize_phone_number(update.message.contact.phone_number)
        user_activity[user.username]["phone_number"] = phone_number
        with open(USER_ACTIVITY_FILE, "w") as file:
            json.dump(user_activity, file)

        if phone_number in ALLOWED_NUMBERS:
            context.user_data['verified'] = True
            return await show_main_menu(update, context)
        else:
            await update.message.reply_text("Access denied. Your phone number is not authorized.")
            return ConversationHandler.END
    else:
        keyboard = [[KeyboardButton("Share Phone Number", request_contact=True)]]
        reply_markup = ReplyKeyboardMarkup(keyboard, one_time_keyboard=True, resize_keyboard=True)
        await update.message.reply_text("Please share your phone number using the button below.", reply_markup=reply_markup)
        return ASK_PHONE

async def restart(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Перезапуск процесса. Пожалуйста, отправьте мне Excel файл, который вы хотите обработать.")
    context.user_data.clear()  # Clear previous data to start fresh
    return ASK_FILE  # Directly transition to ASK_FILE state

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    logger.info(f"Cancel command received from user: {update.message.chat.username}")
    await update.message.reply_text("Процесс отменен. Вы можете начать заново, набрав /start.")
    context.user_data.clear()  # Clear any data in case they want to start again
    return ConversationHandler.END

async def show_main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    keyboard = [["Product-Order", "Admin Panel"], ["Help"]]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("Choose a service:", reply_markup=reply_markup)
    return MAIN_MENU

async def main_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text

    if text == "Product-Order":
        await update.message.reply_text("Please upload your Excel file.")
        return PRODUCT_ORDER
    elif text == "Admin Panel":
        if context.user_data.get('verified') and ADMIN_PHONE in ALLOWED_NUMBERS:
            return await admin_panel(update, context)
        else:
            await update.message.reply_text("Access denied. Admin only.")
            return MAIN_MENU
    elif text == "Help":
        await update.message.reply_text("Usage instructions:\n1. Share phone number.\n2. Choose service.\n3. Follow prompts.")
        return HELP_MENU
    
async def admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    keyboard = [["View Allowed List", "Add Phone"], ["Remove Phone", "User Activity"], ["Back"]]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("Admin Panel:", reply_markup=reply_markup)
    return ADMIN_PANEL


async def handle_admin_action(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    text = update.message.text

    if text == "View Allowed List":
        allowed_list = "\n".join(ALLOWED_NUMBERS)
        await update.message.reply_text(f"Allowed Numbers:\n{allowed_list}")
    elif text == "Add Phone":
        await update.message.reply_text("Send the phone number to add:")
        context.user_data['admin_action'] = "add_phone"
        return ADMIN_PANEL
    elif text == "Remove Phone":
        await update.message.reply_text("Send the phone number to remove:")
        context.user_data['admin_action'] = "remove_phone"
        return ADMIN_PANEL
    elif text == "User Activity":
        output = BytesIO()
        df = pd.DataFrame.from_dict(user_activity, orient="index")
        df.to_excel(output, index=True)
        output.seek(0)
        await update.message.reply_document(document=output, filename="user_activity.xlsx")
    elif text == "Back":
        return await show_main_menu(update, context)
    return ADMIN_PANEL


async def handle_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    try:
        user = update.message.from_user
        document = update.message.document
        file = await document.get_file()
        file_data = BytesIO()
        await file.download_to_memory(file_data)
        file_data.seek(0)

        # Process the Excel file (dummy implementation for now)
        df = pd.read_excel(file_data)
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

        await update.message.reply_document(document=output, filename="processed_data.xlsx", caption="Here is your processed file.")
        return await show_main_menu(update, context)
    except Exception as e:
        await update.message.reply_text("An error occurred while processing the file.")
        return PRODUCT_ORDER

async def back_to_main(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    return await show_main_menu(update, context)

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
            ASK_PHONE: [MessageHandler(filters.CONTACT, handle_phone)],
            MAIN_MENU: [MessageHandler(filters.TEXT & ~filters.COMMAND, main_menu_handler)],
            PRODUCT_ORDER: [MessageHandler(filters.Document.FileExtension("xlsx"), handle_file)],
            ADMIN_PANEL: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_admin_action)],
        },
        fallbacks=[CommandHandler("cancel", cancel), MessageHandler(filters.TEXT & ~filters.COMMAND, back_to_main)]  # Adding fallbacks for /cancel and /restart
    )

    application.add_handler(conv_handler)

    # Run the bot using long polling
    application.run_polling()

if __name__ == "__main__":
    main()
