import os
import json
import logging
import io

from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ConversationHandler, ContextTypes
)
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
import openpyxl

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ['BOT_TOKEN']
SHEET_ID = os.environ['SHEET_ID']
DRIVE_FOLDER_ID = os.environ['DRIVE_FOLDER_ID']
RATE = 1450

WAITING_SUPPLIER = 1

PRICE_KEYWORDS = ['supply', 'supply price', '공급가', '납품가', '공급']
NAME_KEYWORDS = ['product name', 'product', '상품명', '상품', 'item']
EXCLUDE_PRICE = ['retail', 'msrp', 'recommend', 'consumer']


def get_creds():
    creds_json = os.environ['GOOGLE_CREDENTIALS']
    creds_dict = json.loads(creds_json)
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    return Credentials.from_service_account_info(creds_dict, scopes=scopes)


def find_columns(headers):
    product_col = -1
    price_col = -1
    h_lower = [str(h).lower().strip() for h in headers]

    for i, h in enumerate(h_lower):
        for kw in NAME_KEYWORDS:
            if kw in h:
                product_col = i
                break

    for i, h in enumerate(h_lower):
        has_price = any(kw in h for kw in PRICE_KEYWORDS)
        has_exclude = any(ex in h for ex in EXCLUDE_PRICE)
        if has_price and not has_exclude:
            price_col = i
            break

    return product_col, price_col


def update_sheet(data_rows, supplier, date_str):
    creds = get_creds()
    gc = gspread.authorize(creds)
    sheet = gc.open_by_key(SHEET_ID).sheet1

    existing = sheet.get_all_values()
    if not existing or existing[0] != ['Product', 'Supplier', 'Price KRW', 'Price USD', 'Updated']:
        sheet.insert_row(['Product', 'Supplier', 'Price KRW', 'Price USD', 'Updated'], 1)
        existing = sheet.get_all_values()

    added = 0
    updated = 0

    for product, price_krw in data_rows:
        price_usd = round(price_krw / RATE, 2)
        found = False
        for j, row in enumerate(existing[1:], 2):
            if len(row) >= 2 and row[0] == product and row[1] == supplier:
                sheet.update(f'C{j}:E{j}', [[price_krw, price_usd, date_str]])
                updated += 1
                found = True
                break
        if not found:
            sheet.append_row([product, supplier, price_krw, price_usd, date_str])
            added += 1

    return added, updated


def save_to_drive(file_bytes, file_name, supplier):
    creds = get_creds()
    drive = build('drive', 'v3', credentials=creds)

    query = f"name='{supplier}' and '{DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    results = drive.files().list(q=query).execute()
    folders = results.get('files', [])

    if folders:
        supplier_folder_id = folders[0]['id']
    else:
        meta = {
            'name': supplier,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [DRIVE_FOLDER_ID]
        }
        folder = drive.files().create(body=meta).execute()
        supplier_folder_id = folder['id']

    file_meta = {'name': file_name, 'parents': [supplier_folder_id]}
    media = MediaIoBaseUpload(
        io.BytesIO(file_bytes),
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    drive.files().create(body=file_meta, media_body=media).execute()


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        '👋 Привет! Я помогу отслеживать цены поставщиков.\n\n'
        'Просто кидай Excel файл — я спрошу поставщика и добавлю всё в базу.'
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xls')):
        await update.message.reply_text('❌ Нужен Excel файл (.xlsx)')
        return ConversationHandler.END

    context.user_data['file_id'] = doc.file_id
    context.user_data['file_name'] = doc.file_name

    # Check if supplier was sent as caption
    caption = update.message.caption
    if caption and caption.strip():
        context.user_data['supplier'] = caption.strip()
        await update.message.reply_text(f'⏳ Обрабатываю файл от {caption.strip()}...')
        return await process_file(update, context)

    await update.message.reply_text('🏢 Кто поставщик?')
    return WAITING_SUPPLIER


async def handle_supplier_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    supplier = update.message.text.strip()
    context.user_data['supplier'] = supplier
    await update.message.reply_text(f'⏳ Обрабатываю файл от {supplier}...')
    return await process_file(update, context)


async def process_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file_id = context.user_data.get('file_id')
    file_name = context.user_data.get('file_name')
    supplier = context.user_data.get('supplier')

    try:
        file = await context.bot.get_file(file_id)
        file_bytes = await file.download_as_bytearray()

        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        product_col, price_col = find_columns(headers)

        if product_col == -1 or price_col == -1:
            cols = ', '.join([str(h) for h in headers[:15] if h])
            await update.message.reply_text(
                f'❌ Не нашёл столбцы товара или цены.\n'
                f'Столбцы в файле: {cols}\n\n'
                f'Напиши какой столбец цена и какой название.'
            )
            context.user_data.clear()
            return ConversationHandler.END

        data_rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            product = str(row[product_col] or '').strip()
            try:
                price_krw = int(float(str(row[price_col] or 0).replace(',', '').replace(' ', '')))
            except Exception:
                price_krw = 0
            if product and price_krw > 0:
                data_rows.append((product, price_krw))

        if not data_rows:
            await update.message.reply_text('❌ Не нашёл данных в файле.')
            context.user_data.clear()
            return ConversationHandler.END

        date_str = str(update.message.date.strftime('%Y-%m-%d'))
        added, updated_count = update_sheet(data_rows, supplier, date_str)
        save_to_drive(bytes(file_bytes), file_name, supplier)

        await update.message.reply_text(
            f'✅ Готово!\n\n'
            f'📦 Поставщик: {supplier}\n'
            f'➕ Добавлено: {added} товаров\n'
            f'🔄 Обновлено: {updated_count} товаров\n'
            f'📁 Файл сохранён в Drive → папка {supplier}'
        )

    except Exception as e:
        logger.error(e, exc_info=True)
        await update.message.reply_text(f'❌ Ошибка: {str(e)}')

    context.user_data.clear()
    return ConversationHandler.END


def main():
    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[MessageHandler(filters.Document.ALL, handle_document)],
        states={
            WAITING_SUPPLIER: [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_supplier_text)]
        },
        fallbacks=[]
    )

    app.add_handler(CommandHandler('start', start))
    app.add_handler(conv)

    port = int(os.environ.get('PORT', 8080))
    webhook_url = os.environ.get('WEBHOOK_URL', '')

    if webhook_url:
        app.run_webhook(
            listen='0.0.0.0',
            port=port,
            webhook_url=f'{webhook_url}/{BOT_TOKEN}',
            url_path=BOT_TOKEN,
        )
    else:
        app.run_polling()


if __name__ == '__main__':
    main()
