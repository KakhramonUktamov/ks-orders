import os
from datetime import datetime
import logging
import pandas as pd
import json
from collections import defaultdict
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, ConversationHandler, CallbackQueryHandler
from io import BytesIO

logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)


# Admin Telegram ID (set as Heroku environment variable)
ADMIN_TELEGRAM_ID = os.environ.get("ADMIN_TELEGRAM_ID")
BOT_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ALLOWED_NUMBERS_FILE = "allowed_numbers.json"
USER_ACTIVITY_FILE = "user_activity.json"


ASK_FILE, ASK_DAYS, ASK_BRAND, ASK_PERCENTAGE, MAIN_MENU, ADMIN_PANEL, ADD_PHONE, REMOVE_PHONE = range(8)  # Define the states



DEFAULT_ADMIN_PHONE = "+998916919534"


# Load allowed numbers
try:
    if os.path.exists(ALLOWED_NUMBERS_FILE):
        with open(ALLOWED_NUMBERS_FILE, "r") as file:
            ALLOWED_NUMBERS = json.load(file)
            if not isinstance(ALLOWED_NUMBERS, list):
                raise ValueError("Invalid data format in allowed_numbers.json")
    else:
        ALLOWED_NUMBERS = []
except (ValueError, FileNotFoundError):
    ALLOWED_NUMBERS = []

if DEFAULT_ADMIN_PHONE not in ALLOWED_NUMBERS:
    ALLOWED_NUMBERS.append(DEFAULT_ADMIN_PHONE)
    with open(ALLOWED_NUMBERS_FILE, "w") as file:
        json.dump(ALLOWED_NUMBERS, file)


# Dictionary to track user activity
user_activity = defaultdict(lambda: {
    "usage_count": 0, 
    "phone_number": None,
    "last_used": None
})

# Load previous activity from file
if os.path.exists(USER_ACTIVITY_FILE):
    with open(USER_ACTIVITY_FILE, "r") as file:
        user_activity.update(json.load(file))

def save_user_activity():
    with open(USER_ACTIVITY_FILE, "w") as file:
        json.dump(user_activity, file)

def save_allowed_numbers():
    """Save allowed numbers to a JSON file"""
    with open(ALLOWED_NUMBERS_FILE, "w") as file:
        json.dump(ALLOWED_NUMBERS, file)


def normalize_phone_number(phone_number: str) -> str:
    """Normalize phone numbers to international format."""
    phone_number = "".join(c for c in phone_number if c.isdigit() or c == "+")
    return "+" + phone_number if not phone_number.startswith("+") else phone_number

# Handlers
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Start the bot and verify phone number."""
    if 'verified' in context.user_data:
        await show_main_menu(update, context)
        return MAIN_MENU

    keyboard = [[KeyboardButton("Share Phone Number", request_contact=True)]]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("Please share your phone number for verification.", reply_markup=reply_markup)
    return ASK_FILE


async def handle_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Verify phone number."""
    user = update.message.from_user
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    if update.message.contact:
        # Handle phone number shared via contact button
        phone_number = normalize_phone_number(update.message.contact.phone_number)
    else:
        # Handle manually entered phone numbers
        phone_number = normalize_phone_number(update.message.text.strip())

    if phone_number in ALLOWED_NUMBERS:
        # Update user activity only for verified users
        user_activity[user.username].update({
            "phone_number": phone_number,
            "usage_count": user_activity[user.username]["usage_count"] + 1,
            "last_used": now
        })
        save_user_activity()

        context.user_data['verified'] = True
        await update.message.reply_text("✅ Your phone number has been verified.")
        await show_main_menu(update, context)
        return MAIN_MENU
    else:
        await update.message.reply_text("❌ Access Denied. Your phone number is not authorized.")
        return ConversationHandler.END




async def show_main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Display the main menu."""
    keyboard = [
        [KeyboardButton("📊 Product-Order"), KeyboardButton("🛠️ Admin Panel")],
        [KeyboardButton("ℹ️ Help")]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("Choose an option:", reply_markup=reply_markup)
    return MAIN_MENU


async def main_menu_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle main menu actions."""
    text = update.message.text
    if text == "📊 Product-Order":
        await update.message.reply_text("Please send the Excel file to process.")
        return ASK_FILE
    elif text == "🛠️ Admin Panel":
        if str(update.message.chat.id) == ADMIN_TELEGRAM_ID:
            await show_admin_panel(update, context)
            return ADMIN_PANEL
        else:
            await update.message.reply_text("You are not authorized to access the Admin Panel.")
            return MAIN_MENU
    elif text == "ℹ️ Help":
        await update.message.reply_text("ℹ️ How to use this bot:\n1. Verify your phone.\n2. Use Product-Order for file processing.")
        return MAIN_MENU


async def show_admin_panel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Display admin panel options."""
    keyboard = [
        [KeyboardButton("📜 Access List"), KeyboardButton("➕ Add Phone"), KeyboardButton("➖ Remove Phone")],
        [KeyboardButton("📊 Activity List"), KeyboardButton("⬅️ Back")]
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True, one_time_keyboard=True)
    await update.message.reply_text("Admin Panel Options:", reply_markup=reply_markup)
    return ADMIN_PANEL


async def admin_panel_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle admin panel actions."""
    text = update.message.text
    if text == "📜 Access List":
        access_list = "\n".join(ALLOWED_NUMBERS) or "No numbers authorized."
        await update.message.reply_text(f"📜 Access List:\n{access_list}")
    elif text == "➕ Add Phone":
        await update.message.reply_text("Send the phone number to add:")
        return ADD_PHONE
    elif text == "➖ Remove Phone":
        await update.message.reply_text("Send the phone number to remove:")
        return REMOVE_PHONE
    elif text == "📊 Activity List":
        await download_activity_list(update, context)
    elif text == "⬅️ Back":
        await show_main_menu(update, context)
        return MAIN_MENU
    return ADMIN_PANEL


async def add_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Add a phone number to the allowed list."""
    phone_number = normalize_phone_number(update.message.text)
    if phone_number not in ALLOWED_NUMBERS:
        ALLOWED_NUMBERS.append(phone_number)
        save_allowed_numbers()
        await update.message.reply_text(f"✅ {phone_number} added to the access list.")
    else:
        await update.message.reply_text(f"ℹ️ {phone_number} is already in the access list.")
    return ADMIN_PANEL


async def remove_phone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Remove a phone number from the allowed list."""
    phone_number = normalize_phone_number(update.message.text)
    if phone_number in ALLOWED_NUMBERS:
        ALLOWED_NUMBERS.remove(phone_number)
        save_allowed_numbers()
        await update.message.reply_text(f"✅ {phone_number} removed from the access list.")
    else:
        await update.message.reply_text(f"ℹ️ {phone_number} is not in the access list.")
    return ADMIN_PANEL

async def download_activity_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Generate and send the user activity list as an Excel file."""
    if str(update.message.chat.id) != ADMIN_TELEGRAM_ID:
        await update.message.reply_text("❌ You are not authorized to access this feature.")
        return ADMIN_PANEL

    data = [
        {"Username": username, "Usage Count": details["usage_count"], "Phone Number": details["phone_number"], "Last Used": details["last_used"]}
        for username, details in user_activity.items()
    ]
    if not data:
        await update.message.reply_text("No user activity data found.")
        return ADMIN_PANEL

    df = pd.DataFrame(data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="User Activity")
    output.seek(0)

    await update.message.reply_document(
        document=output,
        filename="user_activity.xlsx",
        caption="📊 Here is the user activity list."
    )
    return ADMIN_PANEL


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
    context.user_data.clear()
    user = update.message.from_user
    document = update.message.document
    file_name = document.file_name
    file_size = document.file_size

    logger.info(f"User {user.username} (ID: {user.id}) uploaded file: {file_name} (Size: {file_size} bytes)")

    file = await update.message.document.get_file()
    excel_bytes = BytesIO()
    await file.download_to_memory(excel_bytes)
    excel_bytes.seek(0)
    try:
        # Load the data into a DataFrame and store it
        data = pd.read_excel(excel_bytes)
        context.user_data['data'] = data  # Store the DataFrame for further processing
        logger.info(f"File from {user.username} (ID: {user.id}) is being processed.")
        await update.message.reply_text("Теперь, пожалуйста, введите количество дней для overstock:")
        return ASK_DAYS
    except ValueError as e:
        logger.error(f"Error processing file from {user.username} (ID: {user.id}): {e}")
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


async def stats(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Send user activity logs as an Excel file to the admin, optionally filtered by date range."""
    user_id = str(update.message.chat.id)
    admin_id = str(ADMIN_TELEGRAM_ID)

    # Check if the user is authorized
    if user_id != admin_id:
        await update.message.reply_text("You are not authorized to use this command.")
        return

    # Parse optional date arguments
    args = context.args
    start_date = None
    end_date = None

    try:
        if len(args) >= 2:
            start_date = datetime.strptime(args[0], "%Y-%m-%d")
            end_date = datetime.strptime(args[1], "%Y-%m-%d")
        elif len(args) == 1:
            start_date = datetime.strptime(args[0], "%Y-%m-%d")
            end_date = start_date  # Single date means exact match
        else:
            # No date provided; include all data
            start_date = None
            end_date = None
    except ValueError:
        logger.error("Invalid date format provided in /stats command.")
        await update.message.reply_text(
            "Invalid date format. Please use `/stats YYYY-MM-DD [YYYY-MM-DD]`."
        )
        return

    logger.info(f"Generating stats. Start Date: {start_date}, End Date: {end_date}")

    # Filter user activity based on date range
    filtered_data = []
    for username, details in user_activity.items():
        last_used = details.get("last_used")
        logger.debug(f"Checking user: {username}, Last Used: {last_used}")
        if last_used:
            try:
                last_used_date = datetime.strptime(last_used, "%Y-%m-%d %H:%M:%S")
                if start_date and end_date:
                    # Filter by date range
                    if start_date <= last_used_date <= end_date:
                        filtered_data.append({
                            "Username": username,
                            "Usage Count": details["usage_count"],
                            "Phone Number": details["phone_number"],
                            "Last Used": details["last_used"]
                        })
                else:
                    # Include all data when no date range is specified
                    filtered_data.append({
                        "Username": username,
                        "Usage Count": details["usage_count"],
                        "Phone Number": details["phone_number"],
                        "Last Used": details["last_used"]
                    })
            except ValueError:
                logger.error(f"Invalid date format in last_used for user {username}: {last_used}")
                continue

    if not filtered_data:
        await update.message.reply_text("No activity found.")
        return

    logger.info(f"Filtered Data: {filtered_data}")

    # Convert filtered data to a DataFrame
    df = pd.DataFrame(filtered_data)

    # Generate an Excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="User Activity")
    output.seek(0)

    # Send the Excel file
    await update.message.reply_document(
        document=output,
        filename="user_activity.xlsx" if not start_date else "user_activity_filtered.xlsx",
        caption=f"📊 User activity log{' from all time' if not start_date else f' from {start_date.date()} to {end_date.date()}'}."
    )




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
            ASK_FILE: [MessageHandler(filters.Document.FileExtension("xlsx"), handle_file)],
            ASK_DAYS: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_days)],
            ASK_BRAND: [CallbackQueryHandler(handle_brand)],
            ASK_PERCENTAGE: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_percentage)], 
            MAIN_MENU: [MessageHandler(filters.TEXT & ~filters.COMMAND, main_menu_handler)],
            ADMIN_PANEL: [MessageHandler(filters.TEXT & ~filters.COMMAND, admin_panel_handler)],
            ADD_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, add_phone)],
            REMOVE_PHONE: [MessageHandler(filters.TEXT & ~filters.COMMAND, remove_phone)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],  # Adding fallbacks for /cancel and /restart
    )

    application.add_handler(conv_handler)
    # Run the bot using long polling
    application.run_polling()

if __name__ == "__main__":
    main()
