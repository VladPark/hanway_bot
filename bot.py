import os
import json
import logging
import io
import base64
import threading
import re
from http.server import HTTPServer, BaseHTTPRequestHandler

from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    filters, ConversationHandler, ContextTypes
)
import gspread
import openpyxl
import xlrd

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.environ['BOT_TOKEN']
SHEET_ID = os.environ['SHEET_ID']
CHANNEL_ID = int(os.environ.get('CHANNEL_ID', '0'))
RATE = 1450

WAITING_SUPPLIER = 1

# Sheet schema: ['Дата', 'Поставщик', 'Товар', 'Цена KRW', 'Цена USD']
HEADER = ['Дата', 'Поставщик', 'Товар', 'Цена KRW', 'Цена USD']

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

VOLUME_KEYWORDS = [
    '용량', '내용량', 'volume', 'capacity', 'size', '사이즈', '규격', '중량',
    'weight', 'ml', 'g(', '(g)', 'oz', 'liter'
]
QTY_KEYWORDS = [
    '박스수량', '입수', 'box qty', 'pcs/box', 'qty/box', '개입', '박스당',
    'units/box', 'pieces/box', 'ea/box', 'per box'
]


class HealthHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b'OK')

    def log_message(self, *args):
        pass


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


def get_base_sheet(spreadsheet):
    """Get or create 'База' sheet. Migrates old format automatically."""
    sheet = None
    try:
        sheet = spreadsheet.worksheet('База')
    except gspread.exceptions.WorksheetNotFound:
        sheet = spreadsheet.sheet1
        try:
            sheet.update_title('База')
        except Exception:
            pass

    existing = sheet.get_all_values()

    # Migrate old format: ['Product', 'Supplier', 'Price KRW', 'Price USD', 'Updated']
    if existing and existing[0] == ['Product', 'Supplier', 'Price KRW', 'Price USD', 'Updated']:
        logger.info('Migrating sheet to new format...')
        new_data = [HEADER]
        for row in existing[1:]:
            if len(row) >= 5 and row[0] and row[2]:
                # old: product, supplier, price_krw, price_usd, date
                new_data.append([row[4], row[1], row[0], row[2], row[3]])
        sheet.clear()
        if len(new_data) > 1:
            sheet.update(new_data, value_input_option='RAW')
        else:
            sheet.append_row(HEADER)
    elif not existing or existing[0] != HEADER:
        sheet.clear()
        sheet.append_row(HEADER)

    return sheet


def read_excel(file_bytes, file_name):
    ext = file_name.lower().rsplit('.', 1)[-1]

    if ext == 'xls':
        wb = xlrd.open_workbook(file_contents=bytes(file_bytes))
        best = max(range(wb.nsheets), key=lambda i: wb.sheet_by_index(i).nrows * wb.sheet_by_index(i).ncols)
        ws = wb.sheet_by_index(best)
        all_rows = [[ws.cell_value(r, c) for c in range(ws.ncols)] for r in range(ws.nrows)]
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        best_ws = wb.active
        for s in wb.worksheets:
            if s.max_row and s.max_column:
                if not best_ws.max_row or (s.max_row * s.max_column > best_ws.max_row * best_ws.max_column):
                    best_ws = s
        all_rows = [list(row) for row in best_ws.iter_rows(values_only=True)]
        wb.close()

    if not all_rows:
        return [], []

    header_idx = 0
    for i, row in enumerate(all_rows[:15]):
        if sum(1 for v in row if v is not None and str(v).strip()) >= 3:
            header_idx = i
            break

    return all_rows[header_idx], all_rows[header_idx + 1:]


def find_columns(headers, sample_rows=None):
    product_col = price_col = volume_col = qty_col = -1
    h_lower = [str(h).lower().strip().replace('\n', ' ').replace('\r', ' ') if h else '' for h in headers]

    # First pass: prefer ENG product columns
    for i, h in enumerate(h_lower):
        for kw in NAME_KEYWORDS:
            if kw in h and ('eng' in h or 'english' in h or '(en' in h):
                product_col = i
                break

    # Second pass: any name column
    if product_col == -1:
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

    for i, h in enumerate(h_lower):
        if i in (product_col, price_col):
            continue
        if any(kw in h for kw in VOLUME_KEYWORDS):
            volume_col = i
            break

    for i, h in enumerate(h_lower):
        if i in (product_col, price_col, volume_col):
            continue
        if any(kw in h for kw in QTY_KEYWORDS):
            qty_col = i
            break

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

    return product_col, price_col, volume_col, qty_col


def _extract_unit_from_header(header_str):
    h = header_str.lower()
    for unit in ['ml', 'g', 'oz', 'kg', 'mg', 'l']:
        if unit in h:
            return unit
    return ''


def _format_spec(val, header=''):
    s = str(val).strip()
    if not s or s in ('0', '0.0', 'None', ''):
        return ''
    if any(c.isalpha() for c in s):
        return s
    unit = _extract_unit_from_header(header)
    return s + unit if unit else s


def get_latest_prices(all_data):
    """
    From append-only history, return the most recent price per (product, supplier).
    Returns: dict  (product, supplier) -> (date_str, price_krw)
    """
    latest = {}
    for row in all_data[1:]:           # skip header
        if len(row) < 4:
            continue
        date, supplier, product, price_raw = row[0], row[1], row[2], row[3]
        if not product or not supplier or not price_raw:
            continue
        try:
            price_krw = int(float(str(price_raw).replace(',', '')))
        except (ValueError, TypeError):
            continue
        if price_krw <= 0:
            continue
        key = (product, supplier)
        # Date is YYYY-MM-DD — string comparison works correctly
        if key not in latest or date > latest[key][0]:
            latest[key] = (date, price_krw)
    return latest


def update_sheet(data_rows, supplier, date_str, spreadsheet):
    """Append new price rows to history. Never overwrites existing data."""
    sheet = get_base_sheet(spreadsheet)

    new_rows = [
        [date_str, supplier, product, price_krw, round(price_krw / RATE, 2)]
        for product, price_krw in data_rows
    ]
    sheet.append_rows(new_rows, value_input_option='RAW')

    try:
        rebuild_comparison(spreadsheet)
    except Exception as e:
        logger.warning('Comparison rebuild failed: %s', e)

    return len(new_rows)


def rebuild_comparison(spreadsheet):
    """Rebuild 'Сравнение' pivot sheet from latest prices."""
    try:
        base_sheet = spreadsheet.worksheet('База')
    except Exception:
        base_sheet = spreadsheet.sheet1

    all_data = base_sheet.get_all_values()
    if len(all_data) <= 1:
        return

    latest = get_latest_prices(all_data)

    # pivot: product -> {supplier: price_krw}
    pivot = {}
    suppliers_set = set()
    for (product, supplier), (_, price_krw) in latest.items():
        pivot.setdefault(product, {})[supplier] = price_krw
        suppliers_set.add(supplier)

    if not pivot:
        return

    suppliers = sorted(suppliers_set)

    try:
        comp = spreadsheet.worksheet('Сравнение')
    except gspread.exceptions.WorksheetNotFound:
        comp = spreadsheet.add_worksheet(
            title='Сравнение', rows=len(pivot) + 10, cols=len(suppliers) + 4
        )

    header = ['Товар'] + suppliers + ['Лучший поставщик', 'Мин. цена ₩', 'Мин. цена $']
    table = [header]

    for product in sorted(pivot.keys()):
        prices = pivot[product]
        row = [product] + [prices.get(s, '') for s in suppliers]
        best = min(prices, key=prices.get)
        min_krw = min(prices.values())
        row.extend([best, min_krw, round(min_krw / RATE, 2)])
        table.append(row)

    comp.clear()
    comp.update(table, value_input_option='RAW')


def _esc(s):
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _split_spec(product_name):
    m = re.match(r'^(.+?)\s*\(([^（）()]{1,30})\)\s*$', product_name)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return product_name, ''


async def handle_query(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    if not text or len(text) < 2:
        return

    try:
        gc = get_gspread_client()
        spreadsheet = gc.open_by_key(SHEET_ID)
        try:
            sheet = spreadsheet.worksheet('База')
        except Exception:
            sheet = spreadsheet.sheet1
        all_data = sheet.get_all_values()
    except Exception as e:
        await update.message.reply_text(f'❌ Ошибка чтения базы: {e}')
        return

    if len(all_data) <= 1:
        await update.message.reply_text('База пустая — сначала загрузи файлы поставщиков.')
        return

    text_lower = text.lower()
    latest = get_latest_prices(all_data)

    # Group matching entries by base product name
    products = {}  # base_name -> [(supplier, price_krw, spec, date)]
    for (product, supplier), (date, price_krw) in latest.items():
        if text_lower not in product.lower() and text_lower not in supplier.lower():
            continue
        base_name, spec = _split_spec(product)
        products.setdefault(base_name, []).append((supplier, price_krw, spec, date))

    if not products:
        await update.message.reply_text(
            f'Ничего не найдено по запросу «{_esc(text)}».\n'
            f'Попробуй часть названия бренда, товара или поставщика.'
        )
        return

    lines = [f'🔍 <b>{_esc(text)}</b> — {len(products)} позиций\n']
    shown = 0

    for base_name in sorted(products.keys()):
        if shown >= 20:
            lines.append(f'<i>...и ещё {len(products) - shown} товаров. Уточни запрос.</i>')
            break

        # Sort by price ascending — cheapest first
        entries = sorted(products[base_name], key=lambda x: x[1])
        lines.append(f'<b>{_esc(base_name)}</b>')

        for i, (supplier, price_krw, spec, date) in enumerate(entries):
            price_usd = round(price_krw / RATE, 2)
            best_mark = ' ✅' if i == 0 else ''
            spec_str = f'  <i>{_esc(spec)}</i>' if spec else ''
            lines.append(
                f'  🏭 {_esc(supplier)}: <b>{price_krw:,}₩</b> / ${price_usd}{best_mark}{spec_str}'
            )

        lines.append('')
        shown += 1

    try:
        await update.message.reply_text('\n'.join(lines), parse_mode='HTML')
    except Exception:
        plain = re.sub(r'<[^>]+>', '', '\n'.join(lines))
        await update.message.reply_text(plain)


async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        '👋 Привет!\n\n'
        '📂 <b>Загрузка прайса:</b> кидай Excel файл поставщика.\n'
        'Укажи поставщика в подписи или я спрошу.\n'
        'Каждая загрузка <b>добавляется в историю</b> с датой — старые цены не удаляются.\n\n'
        '🔍 <b>Поиск цен:</b> напиши название бренда или товара.\n'
        'Покажу актуальные цены от всех поставщиков,\n'
        'самая дешёвая отмечена <b>✅</b> с именем поставщика.\n\n'
        '📊 Лист <b>«Сравнение»</b> в таблице обновляется автоматически.',
        parse_mode='HTML'
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if not doc.file_name.lower().endswith(('.xlsx', '.xls')):
        await update.message.reply_text('❌ Нужен Excel файл (.xlsx или .xls)')
        return ConversationHandler.END

    context.user_data['file_id'] = doc.file_id
    context.user_data['file_name'] = doc.file_name

    caption = (update.message.caption or '').strip()
    if caption:
        context.user_data['supplier'] = caption
        await update.message.reply_text(f'⏳ Обрабатываю файл от <b>{_esc(caption)}</b>...', parse_mode='HTML')
        return await process_file(update, context)

    await update.message.reply_text('🏢 Кто поставщик? (напиши название)')
    return WAITING_SUPPLIER


async def handle_supplier_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    supplier = update.message.text.strip()
    context.user_data['supplier'] = supplier
    await update.message.reply_text(f'⏳ Обрабатываю файл от <b>{_esc(supplier)}</b>...', parse_mode='HTML')
    return await process_file(update, context)


async def process_file(update: Update, context: ContextTypes.DEFAULT_TYPE):
    file_id = context.user_data.get('file_id')
    file_name = context.user_data.get('file_name')
    supplier = context.user_data.get('supplier')

    try:
        tg_file = await context.bot.get_file(file_id)
        file_bytes = await tg_file.download_as_bytearray()

        headers, all_rows = read_excel(file_bytes, file_name)

        if not headers:
            await update.message.reply_text('❌ Файл пустой или не читается.')
            context.user_data.clear()
            return ConversationHandler.END

        product_col, price_col, volume_col, qty_col = find_columns(headers, all_rows)

        if product_col == -1 or price_col == -1:
            cols = ', '.join(str(h) for h in headers[:20] if h)
            await update.message.reply_text(
                f'❌ Не нашёл столбцы товара или цены.\n'
                f'Столбцы в файле: {cols}'
            )
            context.user_data.clear()
            return ConversationHandler.END

        vol_header = str(headers[volume_col]) if volume_col != -1 else ''
        qty_header = str(headers[qty_col]) if qty_col != -1 else ''

        data_rows = []
        for row in all_rows:
            if len(row) <= max(product_col, price_col):
                continue
            product = str(row[product_col] or '').strip()
            try:
                price_str = (
                    str(row[price_col] or 0)
                    .replace(',', '').replace(' ', '')
                    .replace('$', '').replace('₩', '')
                    .replace('¥', '').replace('₺', '').strip()
                )
                price_krw = int(float(price_str))
            except Exception:
                price_krw = 0

            if not product or price_krw <= 0:
                continue

            specs = []
            if volume_col != -1 and volume_col < len(row):
                vol = _format_spec(row[volume_col] or '', vol_header)
                if vol:
                    specs.append(vol)
            if qty_col != -1 and qty_col < len(row):
                qty = _format_spec(row[qty_col] or '', qty_header)
                if qty:
                    specs.append(f'{qty}개/박스')
            if specs:
                product = f'{product} ({", ".join(specs)})'

            data_rows.append((product, price_krw))

        if not data_rows:
            await update.message.reply_text('❌ Не нашёл данных в файле.')
            context.user_data.clear()
            return ConversationHandler.END

        date_str = update.message.date.strftime('%Y-%m-%d')
        gc = get_gspread_client()
        spreadsheet = gc.open_by_key(SHEET_ID)
        added = update_sheet(data_rows, supplier, date_str, spreadsheet)

        if CHANNEL_ID:
            try:
                await context.bot.send_document(
                    chat_id=CHANNEL_ID,
                    document=file_id,
                    caption=f'📦 {supplier} | {date_str} | {added} позиций'
                )
            except Exception as e:
                logger.warning('Channel forward failed: %s', e)

        await update.message.reply_text(
            f'✅ <b>Готово!</b>\n\n'
            f'📦 Поставщик: <b>{_esc(supplier)}</b>\n'
            f'📅 Дата: {date_str}\n'
            f'➕ Добавлено в историю: <b>{added}</b> позиций\n'
            f'📊 Таблица сравнения обновлена',
            parse_mode='HTML'
        )

    except Exception as e:
        logger.error('process_file error', exc_info=True)
        err = str(e) or type(e).__name__
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

    app.add_handler(CommandHandler('start', cmd_start))
    app.add_handler(conv)
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_query))

    port = int(os.environ.get('PORT', 8080))
    health = HTTPServer(('0.0.0.0', port), HealthHandler)
    threading.Thread(target=health.serve_forever, daemon=True).start()

    logger.info('Bot started')
    app.run_polling()


if __name__ == '__main__':
    main()
