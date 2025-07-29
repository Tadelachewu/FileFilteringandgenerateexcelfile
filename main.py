import os
import logging
import pandas as pd
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder, CommandHandler, MessageHandler,
    filters, CallbackContext, CallbackQueryHandler
)
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO
from dotenv import load_dotenv

load_dotenv()
BOT_TOKEN = os.getenv("BOT_TOKENdataanalysis")

# Logging setup
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

USER_DATA = {}

async def start(update: Update, context: CallbackContext):
    await update.message.reply_text("👋 Please send me an Excel (.xlsx) file to begin.")

async def handle_file(update: Update, context: CallbackContext):
    document = update.message.document
    if not document.file_name.endswith(".xlsx"):
        await update.message.reply_text("❌ Please upload a valid Excel (.xlsx) file.")
        return

    file = await context.bot.get_file(document.file_id)
    file_path = f"temp/{document.file_name}"
    os.makedirs("temp", exist_ok=True)
    await file.download_to_drive(file_path)

    df = pd.read_excel(file_path)
    chat_id = update.effective_chat.id
    USER_DATA[chat_id] = {
        "df": df,
        "file_path": file_path,
        "selected_values": []
    }

    keyboard = [
        [InlineKeyboardButton(col, callback_data=f"col_{col}")]
        for col in df.columns
    ]
    await update.message.reply_text(
        "📊 Choose a column to filter data by:",
        reply_markup=InlineKeyboardMarkup(keyboard)
    )

async def handle_column_selection(update: Update, context: CallbackContext):
    query = update.callback_query
    await query.answer()
    col_name = query.data.replace("col_", "")
    chat_id = update.effective_chat.id
    df = USER_DATA[chat_id]["df"]

    USER_DATA[chat_id]["filter_column"] = col_name
    USER_DATA[chat_id]["selected_values"] = []

    unique_vals = df[col_name].dropna().astype(str).unique().tolist()
    keyboard = [
        [InlineKeyboardButton(str(val), callback_data=f"val_{val}")]
        for val in unique_vals[:20]
    ]
    keyboard.append([InlineKeyboardButton("✅ Done", callback_data="val_DONE")])
    reply_markup = InlineKeyboardMarkup(keyboard)

    await query.edit_message_text(
        f"🔍 Choose multiple values from **{col_name}**. Tap ✅ Done when finished:",
        reply_markup=reply_markup
    )

async def handle_filter_value(update: Update, context: CallbackContext):
    query = update.callback_query
    await query.answer()
    chat_id = update.effective_chat.id
    value = query.data.replace("val_", "")

    if value == "DONE":
        col = USER_DATA[chat_id]["filter_column"]
        selected = USER_DATA[chat_id]["selected_values"]
        df = USER_DATA[chat_id]["df"]

        if not selected:
            await query.edit_message_text("⚠️ You haven't selected any values yet.")
            return

        filtered_df = df[df[col].astype(str).isin(selected)]
        USER_DATA[chat_id]["filtered_df"] = filtered_df

        keyboard = [
            [
                InlineKeyboardButton("📄 Excel (.xlsx)", callback_data="format_xlsx"),
                InlineKeyboardButton("📄 CSV", callback_data="format_csv"),
                InlineKeyboardButton("📄 JSON", callback_data="format_json")
            ]
        ]
        await query.edit_message_text(
            f"✅ Filtered {len(filtered_df)} rows by `{col}`.\nChoose a format to download:",
            reply_markup=InlineKeyboardMarkup(keyboard)
        )
        return

    # Accumulate selected values
    if value not in USER_DATA[chat_id]["selected_values"]:
        USER_DATA[chat_id]["selected_values"].append(value)

    await query.answer(text=f"Added: {value}", show_alert=False)


async def send_filtered_file(update: Update, context: CallbackContext):
    query = update.callback_query
    await query.answer()
    chat_id = update.effective_chat.id
    df = USER_DATA[chat_id]["filtered_df"]
    format = query.data.replace("format_", "")

    bio = BytesIO()

    if format == "xlsx":
        wb = Workbook()
        ws = wb.active
        ws.title = "FilteredData"

        # Write data with formatting
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 1:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal="center")
                else:
                    cell.alignment = Alignment(horizontal="left")

        # Auto-size columns
        for column_cells in ws.columns:
            max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
            col_letter = column_cells[0].column_letter
            ws.column_dimensions[col_letter].width = max_length + 2

        wb.save(bio)
        filename = f"filtered_{chat_id}.xlsx"

    elif format == "csv":
        df.to_csv(bio, index=False)
        filename = f"filtered_{chat_id}.csv"

    elif format == "json":
        df.to_json(bio, orient="records", lines=True)
        filename = f"filtered_{chat_id}.json"

    bio.seek(0)
    await context.bot.send_document(chat_id=chat_id, document=bio, filename=filename)
    await query.edit_message_text("✅ Here’s your filtered file!")

def main():
    from telegram.ext import Defaults
    import asyncio

    ENV = os.getenv("ENV", "development")
    WEBHOOK_URL = os.getenv("WEBHOOK_URL")
    PORT = int(os.getenv("PORT", "8443"))

    app = ApplicationBuilder().token(BOT_TOKEN).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.Document.FileExtension("xlsx"), handle_file))
    app.add_handler(CallbackQueryHandler(handle_column_selection, pattern="^col_"))
    app.add_handler(CallbackQueryHandler(handle_filter_value, pattern="^val_"))
    app.add_handler(CallbackQueryHandler(send_filtered_file, pattern="^format_"))

    if ENV == "production":
        print("🚀 Running in production mode with webhook...")
        app.run_webhook(
            listen="0.0.0.0",
            port=PORT,
            webhook_url=WEBHOOK_URL
        )
    else:
        print("🛠️ Running in development mode with polling...")
        app.run_polling()

if __name__ == "__main__":
    main()
