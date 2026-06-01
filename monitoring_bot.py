"""
Magnum Monitoring Bot — облачная версия
Пользователь присылает xlsx → бот анализирует → отвечает отчётом + Excel
"""

import os
import io
import statistics
import telebot
import requests
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
bot = telebot.TeleBot(BOT_TOKEN)

# ── Загрузка маппинга брендов ────────────────────────────────────────────────
def load_brands():
    brands = {}
    try:
        wb = load_workbook("Маппинг.xlsx", data_only=True)
        for row in wb['Бренды'].iter_rows(min_row=2, max_row=wb['Бренды'].max_row, values_only=True):
            if row[0] and row[1] != 'Закрыто':
                brands[str(row[0]).strip().upper()] = {
                    'km': row[1] or '—', 'brand': row[4] or '—'}
    except Exception:
        pass
    return brands

BRANDS = load_brands()

def lookup(name):
    key = name.strip().upper()
    if key in BRANDS: return BRANDS[key]
    for k, v in BRANDS.items():
        if key in k or k in key: return v
    return {'km': '—', 'brand': '—'}

# ── Анализ ───────────────────────────────────────────────────────────────────
def run_analysis(file_bytes):
    wb_src = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws_src = wb_src['Тепловая карта']

    MARKET_IDXS = [4, 5, 6, 8]
    SMALL_IDXS  = [7, 10]
    KASPI_IDX   = 2

    segment, rows = None, []
    for row in ws_src.iter_rows(min_row=8, max_row=65, values_only=True):
        name = row[0]
        if name in ('HARD', 'SOFT', 'СЕЗОН'):
            segment = name; continue
        if not name or name == '(пусто)': continue

        magnum = row[1]
        kaspi  = row[KASPI_IDX]
        market = [row[i] for i in MARKET_IDXS if row[i] is not None]
        smalls = [row[i] for i in SMALL_IDXS  if row[i] is not None]
        if smalls: market.append(round(statistics.mean(smalls)))
        if not market: continue

        avg  = round(statistics.mean(market))
        mn   = min(market)
        info = lookup(name)

        if magnum:
            pct = round((magnum - avg) / avg * 100, 1)
            if pct >= 15:    status = 'РИСК'
            elif pct >= 0:   status = 'Дороже'
            elif pct >= -5:  status = 'Норма'
            else:            status = 'Дешевле'
        else:
            pct, status = None, 'Нет цены'

        rows.append(dict(segment=segment, name=name, magnum=magnum, kaspi=kaspi,
                         avg=avg, min=mn, pct=pct, status=status,
                         km=info['km'], brand=info['brand']))
    return rows

# ── Генерация Excel ──────────────────────────────────────────────────────────
def build_excel(rows):
    FN = 'Arial'
    HDR_FILL = PatternFill('solid', start_color='1F3864')
    HDR_FONT = Font(name=FN, bold=True, color='FFFFFF', size=10)
    SEG_FILLS = {'HARD':  PatternFill('solid', start_color='C6EFCE'),
                 'SOFT':  PatternFill('solid', start_color='DDEBF7'),
                 'СЕЗОН': PatternFill('solid', start_color='FFF2CC')}
    SF = {'РИСК':     PatternFill('solid', start_color='FFC7CE'),
          'Дороже':   PatternFill('solid', start_color='FFEB9C'),
          'Норма':    PatternFill('solid', start_color='FFFFFF'),
          'Дешевле':  PatternFill('solid', start_color='C6EFCE'),
          'Нет цены': PatternFill('solid', start_color='EDEDED')}
    SFONTS = {'РИСК':     Font(name=FN, color='9C0006', bold=True, size=10),
              'Дороже':   Font(name=FN, color='9C6500', size=10),
              'Норма':    Font(name=FN, size=10),
              'Дешевле':  Font(name=FN, color='276221', size=10),
              'Нет цены': Font(name=FN, color='7F7F7F', size=10)}
    NORM = Font(name=FN, size=10)
    BOLD = Font(name=FN, bold=True, size=10)
    thin = Side(style='thin', color='BFBFBF')
    BRD  = Border(left=thin, right=thin, top=thin, bottom=thin)

    def cs(ws, r, c, val, fill=None, font=None, align='left', num_fmt=None, bold=False):
        cell = ws.cell(row=r, column=c, value=val)
        if fill:   cell.fill = fill
        cell.font      = font or (BOLD if bold else NORM)
        cell.alignment = Alignment(horizontal=align, vertical='center')
        cell.border    = BRD
        if num_fmt: cell.number_format = num_fmt

    wb  = Workbook()
    ws1 = wb.active
    ws1.title = 'Тепловая карта'
    ws1.freeze_panes = 'B3'
    date_str = datetime.now().strftime('%d.%m.%Y')

    HEADERS = ['Товар','Сегмент','Магнум','Ср.рынок','Откл, %','Статус','Каспий','Мин','КМ','Бренд']
    WIDTHS  = [42, 8, 9, 9, 10, 11, 9, 9, 22, 18]

    ws1.merge_cells('A1:J1')
    t = ws1['A1']
    t.value     = f'Мониторинг цен vs конкуренты — {date_str}'
    t.font      = Font(name=FN, bold=True, size=13, color='1F3864')
    t.alignment = Alignment(horizontal='center', vertical='center')
    ws1.row_dimensions[1].height = 22

    for ci, (h, w) in enumerate(zip(HEADERS, WIDTHS), 1):
        cs(ws1, 2, ci, h, fill=HDR_FILL, font=HDR_FONT, align='center')
        ws1.column_dimensions[get_column_letter(ci)].width = w
    ws1.row_dimensions[2].height = 18

    r, prev_seg = 3, None
    for row in rows:
        seg, status = row['segment'], row['status']
        if seg != prev_seg:
            ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
            sc = ws1.cell(row=r, column=1, value=f'  {seg}')
            sc.fill = SEG_FILLS[seg]
            sc.font = Font(name=FN, bold=True, size=10)
            sc.alignment = Alignment(horizontal='left', vertical='center')
            sc.border = BRD
            ws1.row_dimensions[r].height = 16
            r += 1
            prev_seg = seg

        pct_val = (row['pct'] / 100) if row['pct'] is not None else None
        cs(ws1, r, 1,  row['name'],   font=SFONTS[status])
        cs(ws1, r, 2,  seg,           align='center')
        cs(ws1, r, 3,  row['magnum'], fill=SF[status], font=SFONTS[status], align='right', num_fmt='#,##0')
        cs(ws1, r, 4,  row['avg'],    align='right', num_fmt='#,##0')
        cs(ws1, r, 5,  pct_val,       fill=SF[status], font=SFONTS[status], align='right', num_fmt='+0.0%;-0.0%;"-"')
        cs(ws1, r, 6,  status,        fill=SF[status], font=SFONTS[status], align='center')
        cs(ws1, r, 7,  row['kaspi'],  align='right', num_fmt='#,##0')
        cs(ws1, r, 8,  row['min'],    align='right', num_fmt='#,##0')
        cs(ws1, r, 9,  row['km'])
        cs(ws1, r, 10, row['brand'])
        ws1.row_dimensions[r].height = 15
        r += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── Текстовый отчёт ──────────────────────────────────────────────────────────
def format_report(rows, filename):
    date_str  = datetime.now().strftime('%d.%m.%Y')
    risk_rows = sorted(
        [r for r in rows if r['status'] == 'РИСК' and r.get('pct') and r['name'] != 'АВОКАДО КГ'],
        key=lambda x: -x['pct'])
    counts = {}
    for r in rows:
        counts[r['status']] = counts.get(r['status'], 0) + 1

    lines = [
        f"📊 *Мониторинг цен — {date_str}*",
        f"Файл: `{filename}`\n",
        f"🔴 Риск: {counts.get('РИСК', 0)}  "
        f"🟡 Дороже: {counts.get('Дороже', 0)}  "
        f"⚪ Норма: {counts.get('Норма', 0)}  "
        f"🟢 Дешевле: {counts.get('Дешевле', 0)}\n",
        "*ТОП рисков:*"
    ]
    for i, r in enumerate(risk_rows[:10], 1):
        kaspi_str = f" | Каспий: {r['kaspi']:,}" if r['kaspi'] else ""
        lines.append(
            f"{i}\\. {r['name']}\n"
            f"   Магнум: *{r['magnum']:,}* | Рынок: {r['avg']:,}{kaspi_str} | *\\+{r['pct']}%*\n"
            f"   КМ: {r['km']}"
        )
    return "\n".join(lines)

# ── Хендлеры ─────────────────────────────────────────────────────────────────
@bot.message_handler(content_types=['document'])
def handle_document(message):
    doc = message.document
    if not doc.file_name.endswith('.xlsx'):
        bot.reply_to(message, "Пришлите файл в формате .xlsx")
        return

    bot.reply_to(message, "⏳ Анализирую...")

    try:
        file_info = bot.get_file(doc.file_id)
        url = f"https://api.telegram.org/file/bot{BOT_TOKEN}/{file_info.file_path}"
        file_bytes = requests.get(url).content

        rows    = run_analysis(file_bytes)
        report  = format_report(rows, doc.file_name)
        excel   = build_excel(rows)
        date_str = datetime.now().strftime('%d.%m')

        bot.send_message(message.chat.id, report, parse_mode='MarkdownV2')
        bot.send_document(message.chat.id, excel,
                          visible_file_name=f"Анализ мониторинга {date_str}.xlsx",
                          caption=f"📎 Полный анализ — {date_str}")
    except Exception as e:
        bot.reply_to(message, f"❌ Ошибка: {e}")

@bot.message_handler(func=lambda m: True)
def handle_text(message):
    bot.reply_to(message, "Пришлите файл мониторинга в формате .xlsx 📎")

print("✅ Бот запущен (облачный режим)")
bot.infinity_polling()
