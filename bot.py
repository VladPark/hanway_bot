import os
import json
import logging
import io
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

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ['BOT_TOKEN']
SHEET_ID = os.environ['SHEET_ID']
DRIVE_FOLDER_ID = os.environ['DRIVE_FOLDER_ID']
RATE = 1450

WAITING_SUPPLIER = 1

PRICE_KEYWORDS = [
    'supply price', 'supply', '\uacf5\uae09\uac00', '\ub0a9\ud488\uac00', '\uacf5\uae09\ub2e8\uac00', '\uacf5\uae09\uac00\uaca9',
    '\ub2e8\uac00', '\uc6d0\uac00', '\ub3c4\ub9e4\uac00', '\ub3c4\ub9e4\ub2e8\uac00', 'wholesale', 'cost', 'unit price',
    '\uac00\uaca9', '\uacf5\uae09', '\ub9e4\uc785\uac00', '\ub9e4\uc785\ub2e8\uac00', '\uc6d0\ub2e8\uac00', 'price'
]
NAME_KEYWORDS = [
    'product name', 'product', '\uc0c1\ud488\uba85', '\ud488\uba85', '\uc81c\ud488\uba85', '\uc0c1\ud488',
    'item', 'name', '\ud488\ubaa9', '\uc544\uc774\ud15c', '\ubaa8\ub378\uba85', '\ubaa8\ub378', '\uc81c\ud488', '\ud56d\ubaa9'
]
EXCLUDE_PRICE = ['retail', 'msrp', 'recommend', 'consumer', '\uc18c\ube44\uc790', '\ud310\ub9e4\uac00', '\uc18c\ub9e4']


class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b'OK')
    def log_message(self, *args):
        pass


def get_creds():
    creds_json = os.environ['GOOGLE_CREDENTIALS']
    creds_dict = json.loads(creds_json)
    scopes = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
    ]
    return Credentials.from_service_account_info(creds_dict, scopes=scopes)


def find_columns(headers, ws=None):
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

    if ws and (product_col == -1 or price_col == -1):
        numeric_scores = [0] * len(headers)
        text_scores = [0] * len(headers)
        for row in ws.iter_rows(min_row=2, max_row=min(10, ws.max_row), values_only=True):
            for ci, val in enumerate(row):
                if ci >= len(headers):
                    break
                if val is None:
                    continue
                s = str(val).replace(',', '').replace(' ', '').replace('\uc6d0', '')
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
        '\U0001f44b \u041f\u0440\u0438\u0432\u0435\u0442! \u042f \u043f\u043e\u043c\u043e\u0433\u0443 \u043e\u0442\u0441\u043b\u0435\u0436\u0438\u0432\u0430\u0442\u044c \u0446\u0435\u043d\u044b \u043f\u043e\u0441\u0442\u0430\u0432\u0449\u0438\u043a\u043e\u0432.\n\n'
        '\u041f\u0440\u043e\u0441\u0442\u043e \u043a\u0438\u0434\u0430\u0439 Excel \u0444\u0430\u0439\u043b \u2014 \u044f \u0441\u043f\u0440\u043e\u0448\u0443 \u043f\u043e\u0441\u0442\u0430\u0432\u0449\u0438\u043a\u0430 \u0438 \u0434\u043e\u0431\u0430\u0432\u043b\u044e \u0432\u0441\u0451 \u0432 \u0431\u0430\u0437\u0443.'
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xls')):
        await update.message.reply_text('\u274c \u041d\u0443\u0436\u0435\u043d Excel \u0444\u0430\u0439\u043b (.xlsx)')
        return ConversationHandler.END

    context.user_data['file_id'] = doc.file_id
    context.user_data['file_name'] = doc.file_name

    caption = update.message.caption
    if caption and caption.strip():
        context.user_data['supplier'] = caption.strip()
        await update.message.reply_text(f'\u23f3 \u041e\u0431\u0440\u0430\u0431\u0430\u0442\u044b\u0432\u0430\u044e \u0444\u0430\u0439\u043b \u043e\u0442 {caption.strip()}...')
        return await process_file(update, context)

    await update.message.reply_text('\U0001f3e2 \u041a\u0442\u043e \u043f\u043e\u0441\u0442\u0430\u0432\u0449\u0438\u043a?')
    return WAITING_SUPPLIER


async def handle_supplier_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    supplier = update.message.text.strip()
    context.user_data['supplier'] = supplier
    await update.message.reply_text(f'\u23f3 \u041e\u0431\u0440\u0430\u0431\u0430\u0442\u044b\u0432\u0430\u044e \u0444\u0430\u0439\u043b \u043e\u0442 {supplier}...')
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
        product_col, price_col = find_columns(headers, ws)

        if product_col == -1 or price_col == -1:
            cols = ', '.join([str(h) for h in headers[:15] if h])
            await update.message.reply_text(
                f'\u274c \u041d\u0435 \u043d\u0430\u0448\u0451\u043b \u0441\u0442\u043e\u043b\u0431\u0446\u044b \u0442\u043e\u0432\u0430\u0440\u0430 \u0438\u043b\u0438 \u0446\u0435\u043d\u044b.\n'
                f'\u0421\u0442\u043e\u043b\u0431\u0446\u044b \u0432 \u0444\u0430\u0439\u043b\u0435: {cols}\n\n'
                f'\u041d\u0430\u043f\u0438\u0448\u0438 \u043a\u0430\u043a\u043e\u0439 \u0441\u0442\u043e\u043b\u0431\u0435\u0446 \u0446\u0435\u043d\u0430 \u0438 \u043a\u0430\u043a\u043e\u0439 \u043d\u0430\u0437\u0432\u0430\u043d\u0438\u0435.'
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
            await update.message.reply_text('\u274c \u041d\u0435 \u043d\u0430\u0448\u0451\u043b \u0434\u0430\u043d\u043d\u044b\u0445 \u0432 \u0444\u0430\u0439\u043b\u0435.')
            context.user_data.clear()
            return ConversationHandler.END

        date_str = str(update.message.date.strftime('%Y-%m-%d'))
        added, updated_count = update_sheet(data_rows, supplier, date_str)
        save_to_drive(bytes(file_bytes), file_name, supplier)

        await update.message.reply_text(
            f'\u2705 \u0413\u043e\u0442\u043e\u0432\u043e!\n\n'
            f'\U0001f4e6 \u041f\u043e\u0441\u0442\u0430\u0432\u0449\u0438\u043a: {supplier}\n'
            f'\u2795 \u0414\u043e\u0431\u0430\u0432\u043b\u0435\u043d\u043e: {added} \u0442\u043e\u0432\u0430\u0440\u043e\u0432\n'
            f'\U0001f504 \u041e\u0431\u043d\u043e\u0432\u043b\u0435\u043d\u043e: {updated_count} \u0442\u043e\u0432\u0430\u0440\u043e\u0432\n'
            f'\U0001f4c1 \u0424\u0430\u0439\u043b \u0441\u043e\u0445\u0440\u0430\u043d\u0451\u043d \u0432 Drive \u2192 \u043f\u0430\u043f\u043a\u0430 {supplier}'
        )

    except Exception as e:
        logger.error(e, exc_info=True)
        await update.message.reply_text(f'\u274c \u041e\u0448\u0438\u0431\u043a\u0430: {str(e)}')

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
