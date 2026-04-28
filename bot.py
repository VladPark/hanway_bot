import os
import json
import logging
import io
import base64
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler

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
import xlrd

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ['BOT_TOKEN']
SHEET_ID = os.environ['SHEET_ID']
DRIVE_FOLDER_ID = os.environ['DRIVE_FOLDER_ID']
RATE = 1450

WAITING_SUPPLIER = 1

PRICE_KEYWORDS = [
    'supply price', 'supply', '공급가', '납품가', '공급단가', '공급가격',
    '단가', '원가', '도매가', '도매단가', 'wholesale', 'cost', 'unit price',
    '가격', '공급', '매입가', '매입단가', '원단가', 'price'
]
NAME_KEYWORDS = [
    'product name', 'product', '상품명', '품명', '제품명', '상품',
    'item', 'name', '품목', '아이템', '모델명', '모델', '제품', '항목'
]
EXCLUDE_PRICE = ['retail', 'msrp', 'recommend', 'consumer', '소비자', '판매가', '소매']


class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b'OK')
    def log_message(self, *args):
        pass


def get_creds():
    raw = os.environ['GOOGLE_CREDENTIALS']
    try:
        creds_dict = json.loads(raw)
    except (json.JSONDecodeError, ValueError):
        creds_dict = json.loads(base64.b64decode(raw).decode('utf-8'))
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    return Credentials.from_service_account_info(creds_dict, scopes=scopes)


def read_excel(file_bytes, file_name):
    """Returns (headers list, all_data list of lists). Supports .xlsx and .xls."""
    ext = file_name.lower().rsplit('.', 1)[-1]

    if ext == 'xls':
        wb = xlrd.open_workbook(file_contents=bytes(file_bytes))
        # Pick sheet with most cells
        best = max(range(wb.nsheets), key=lambda i: wb.sheet_by_index(i).nrows * wb.sheet_by_index(i).ncols)
        ws = wb.sheet_by_index(best)
        all_rows = [[ws.cell_value(r, c) for c in range(ws.ncols)] for r in range(ws.nrows)]
    else:
        # read_only=True streams row by row — works on large files
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        # Pick sheet with largest declared dimensions
        best_ws = wb.active
        for sheet in wb.worksheets:
            if sheet.max_row and sheet.max_column:
                if not best_ws.max_row or (sheet.max_row * sheet.max_column > best_ws.max_row * best_ws.max_column):
                    best_ws = sheet
        all_rows = [list(row) for row in best_ws.iter_rows(values_only=True)]
        wb.close()

    if not all_rows:
        return [], []

    # Find header row: first row with 3+ non-empty cells
    header_idx = 0
    for i, row in enumerate(all_rows[:15]):
        if sum(1 for v in row if v is not None and str(v).strip()) >= 3:
            header_idx = i
            break

    headers = all_rows[header_idx]
    data = all_rows[header_idx + 1:]
    return headers, data


def find_columns(headers, sample_rows=None):
    product_col = -1
    price_col = -1
    h_lower = [str(h).lower().strip() if h else '' for h in headers]

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

    # Fallback: scan first 10 data rows
    if sample_rows and (product_col == -1 or price_col == -1):
        numeric_scores = [0] * len(headers)
        text_scores = [0] * len(headers)
        for row in sample_rows[:10]:
            for ci, val in enumerate(row):
                if ci >= len(headers) or val is None:
                    continue
                s = str(val).replace(',', '').replace(' ', '').replace('원', '')
                try:
                    float(s)
                    numeric_scores[ci] += 1
                except ValueError:
                    if len(str(val).strip()) > 1:
                        text_scores[ci] += 1

        if price_col == -1:
            best = max(range(len(numeric_scores)), key=lambda i: numeric_scores[i])
            if numeric_scores[best] > 0:
                price_col = best

        if product_col == -1:
            best = max(range(len(text_scores)), key=lambda i: text_scores[i])
            if text_scores[best] > 0 and best != price_col:
                product_col = best

    return product_col, price_col


def get_gspread_client():
    raw = os.environ['GOOGLE_CREDENTIALS']
    try:
        creds_dict = json.loads(raw)
    except (json.JSONDecodeError, ValueError):
        creds_dict = json.loads(base64.b64decode(raw).decode('utf-8'))
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    return gspread.service_account_from_dict(creds_dict, scopes=scopes)


def update_sheet(data_rows, supplier, date_str):
    gc = get_gspread_client()
    sheet = gc.open_by_key(SHEET_ID).sheet1

    existing = sheet.get_all_values()
    if not existing or existing[0] != ['Product', 'Supplier', 'Price KRW', 'Price USD', 'Updated']:
        sheet.insert_row(['Product', 'Supplier', 'Price KRW', 'Price USD', 'Updated'], 1)
        existing = sheet.get_all_values()

    # Build lookup for O(1) access instead of O(n) per product
    existing_lookup = {}
    for j, row in enumerate(existing[1:], 2):
        if len(row) >= 2:
            existing_lookup[(row[0], row[1])] = j

    batch_updates = []
    new_rows = []
    added = 0
    updated = 0

    for product, price_krw in data_rows:
        price_usd = round(price_krw / RATE, 2)
        j = existing_lookup.get((product, supplier))
        if j:
            batch_updates.append({'range': f'C{j}:E{j}', 'values': [[price_krw, price_usd, date_str]]})
            updated += 1
        else:
            new_rows.append([product, supplier, price_krw, price_usd, date_str])
            added += 1

    if batch_updates:
        sheet.batch_update(batch_updates)
    if new_rows:
        sheet.append_rows(new_rows, value_input_option='RAW')

    return added, updated


def save_to_drive(file_bytes, file_name, supplier):
    creds = get_creds()  # google-auth creds for Drive API
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
        await update.message.reply_text('❌ Нужен Excel файл (.xlsx или .xls)')
        return ConversationHandler.END

    context.user_data['file_id'] = doc.file_id
    context.user_data['file_name'] = doc.file_name

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

        headers, all_data = read_excel(file_bytes, file_name)

        if not headers:
            await update.message.reply_text('❌ Файл пустой или не читается.')
            context.user_data.clear()
            return ConversationHandler.END

        product_col, price_col = find_columns(headers, all_data)

        if product_col == -1 or price_col == -1:
            cols = ', '.join([str(h) for h in headers[:20] if h])
            await update.message.reply_text(
                f'❌ Не нашёл столбцы товара или цены.\n'
                f'Столбцы в файле: {cols}\n\n'
                f'Напиши какой столбец цена и какой название.'
            )
            context.user_data.clear()
            return ConversationHandler.END

        data_rows = []
        for row in all_data:
            if len(row) <= max(product_col, price_col):
                continue
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

        drive_note = ''
        try:
            save_to_drive(bytes(file_bytes), file_name, supplier)
            drive_note = f'\n📁 Файл сохранён в Drive → папка {supplier}'
        except Exception as drive_err:
            logger.warning(f'Drive upload failed: {drive_err}')
            drive_note = '\n⚠️ Drive: не удалось сохранить файл (нет квоты у сервисного аккаунта)'

        await update.message.reply_text(
            f'✅ Готово!\n\n'
            f'📦 Поставщик: {supplier}\n'
            f'➕ Добавлено: {added} товаров\n'
            f'🔄 Обновлено: {updated_count} товаров'
            f'{drive_note}'
        )

    except Exception as e:
        logger.error(e, exc_info=True)
        err = str(e) if str(e) else type(e).__name__
        if not str(e) and e.__cause__:
            err = f"{type(e).__name__}: {str(e.__cause__)[:300]}"
        await update.message.reply_text(f'❌ Ошибка: {err}')

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
        health = HTTPServer(('0.0.0.0', port), HealthHandler)
        t = threading.Thread(target=health.serve_forever, daemon=True)
        t.start()
        app.run_polling()


if __name__ == '__main__':
    main()
