import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import io
import math
import calendar
from datetime import date, timedelta
import re

st.set_page_config(page_title="宿泊税計算システム", page_icon="🏨", layout="wide")

TAX_RATE = 0.02
MAX_TAX_PER_PERSON = 2000

def reiwa(year):
    return year - 2018


def parse_daily_csv(file_bytes, target_year, target_month):
    """日次予約CSVを解析し、対象月に宿泊した夜ごとのレコードを返す。
    月をまたぐ滞在は泊数(X/Y)と利用期間から宿泊日を計算し、対象月分のみ抽出。
    Returns: list of {day, guests, amount}
    """
    for enc in ['utf-8-sig', 'shift_jis', 'cp932']:
        try:
            content = file_bytes.decode(enc)
            break
        except Exception:
            continue
    else:
        return []

    items = []

    for line in content.splitlines()[1:]:  # ヘッダー行をスキップ
        parts = [p.strip() for p in line.split(',')]
        if len(parts) < 10:
            continue

        res_no = parts[0]

        # 小計行(YYYY/MM/DD)・合計行(空)・ヘッダー行をスキップ
        if (not res_no
                or res_no == '予約番号'
                or re.match(r'\d{4}/\d{2}/\d{2}', res_no)):
            continue

        # 泊数(col1): "X/Y" → 何泊目か
        m = re.match(r'(\d+)/(\d+)', parts[1])
        if not m:
            continue
        night_num = int(m.group(1))

        # 利用期間(col5): "MM/DD ～ MM/DD" → チェックイン日を取得
        pm = re.match(r'(\d{1,2})/(\d{1,2})', parts[5])
        if not pm:
            continue
        ci_month, ci_day = int(pm.group(1)), int(pm.group(2))

        # チェックイン年の推定:
        # チェックイン月 > 対象月 の場合は前年（例: 12月チェックイン→翌1月が対象月）
        ci_year = target_year - 1 if ci_month > target_month else target_year

        try:
            checkin = date(ci_year, ci_month, ci_day)
            night_date = checkin + timedelta(days=night_num - 1)
        except ValueError:
            continue

        # 対象月以外の泊は除外（月またぎ自動カット）
        if night_date.year != target_year or night_date.month != target_month:
            continue

        # 合計人数(col6) / 基本利用(col9)
        try:
            guests = int(parts[6])
            amt_str = parts[9].replace(',', '')
            amount = int(amt_str) if amt_str.lstrip('-').isdigit() else 0
        except (ValueError, IndexError):
            continue

        if guests <= 0:
            continue

        items.append({'day': night_date.day, 'guests': guests, 'amount': amount})

    return items


def calc_tax(guests, amount, meal_per_person=0, tax_inclusive=False):
    if guests <= 0:
        return 0, 0, 0
    if tax_inclusive:
        amount = amount / 1.10
    net = amount - meal_per_person * guests
    per_person = net / guests
    base = math.floor(per_person / 1000) * 1000
    tax_pp = min(base * TAX_RATE, MAX_TAX_PER_PERSON)
    total_tax = int(tax_pp * guests)
    return int(base), int(base * guests), total_tax


def compute_daily_agg(items, meal_per_person=0, tax_inclusive=False):
    """予約ごとに宿泊税を計算し、日別に集計して返す。
    Returns: {day: {guests, base, tax}}
    """
    daily = {}
    for item in items:
        _, base, tax = calc_tax(
            item['guests'], item['amount'],
            meal_per_person=meal_per_person,
            tax_inclusive=tax_inclusive,
        )
        d = item['day']
        if d not in daily:
            daily[d] = {'guests': 0, 'base': 0, 'tax': 0}
        daily[d]['guests'] += item['guests']
        daily[d]['base'] += base
        daily[d]['tax'] += tax
    return daily


def parse_for_verification(file_bytes, target_year, target_month,
                            meal_per_person=0, tax_inclusive=False):
    """検証用: 予約別の計算明細とCSV小計を返す。
    Returns:
        detail_rows : list[dict]  ─ 予約1件ごとの計算ステップ
        csv_subtotals: {day: {guests, amount}}  ─ CSV 小計行から取得
    """
    for enc in ['utf-8-sig', 'shift_jis', 'cp932']:
        try:
            content = file_bytes.decode(enc)
            break
        except Exception:
            continue
    else:
        return [], {}

    detail_rows = []
    csv_subtotals = {}

    for line in content.splitlines()[1:]:
        parts = [p.strip() for p in line.split(',')]
        if len(parts) < 10:
            continue

        res_no = parts[0]

        # CSV 小計行 (YYYY/MM/DD) を取得
        m_date = re.match(r'(\d{4})/(\d{2})/(\d{2})', res_no)
        if m_date:
            y, mo, d = int(m_date.group(1)), int(m_date.group(2)), int(m_date.group(3))
            if y == target_year and mo == target_month:
                try:
                    csv_subtotals[d] = {
                        'guests': int(parts[6]),
                        'amount': int(parts[9].replace(',', '')),
                    }
                except (ValueError, IndexError):
                    pass
            continue

        if not res_no or res_no == '予約番号':
            continue

        # 泊数
        m_n = re.match(r'(\d+)/(\d+)', parts[1])
        if not m_n:
            continue
        night_num = int(m_n.group(1))

        # 利用期間
        pm = re.match(r'(\d{1,2})/(\d{1,2})', parts[5])
        if not pm:
            continue
        ci_month, ci_day = int(pm.group(1)), int(pm.group(2))
        ci_year = target_year - 1 if ci_month > target_month else target_year

        try:
            checkin = date(ci_year, ci_month, ci_day)
            night_date = checkin + timedelta(days=night_num - 1)
        except ValueError:
            continue

        if night_date.year != target_year or night_date.month != target_month:
            continue

        try:
            guests = int(parts[6])
            amt_str = parts[9].replace(',', '')
            amount = int(amt_str) if amt_str.lstrip('-').isdigit() else 0
        except (ValueError, IndexError):
            continue

        if guests <= 0:
            continue

        # 計算ステップ
        adj_amount = (amount / 1.10) if tax_inclusive else amount
        net = adj_amount - meal_per_person * guests
        per_person_raw = net / guests
        base_per = math.floor(per_person_raw / 1000) * 1000
        tax_pp = min(base_per * TAX_RATE, MAX_TAX_PER_PERSON)
        tax_total = int(tax_pp * guests)

        detail_rows.append({
            '宿泊日': f'{target_month}/{night_date.day}',
            '予約番号': res_no,
            '利用期間': parts[5],
            '泊数': parts[1],
            '氏名': parts[4],
            '人数': guests,
            '基本利用（円）': amount,
            '1人当たり料金': int(per_person_raw),
            '課税標準額/人（千円切捨）': int(base_per),
            '課税標準額合計': int(base_per * guests),
            '宿泊税（円）': tax_total,
        })

    return detail_rows, csv_subtotals


def thin_border():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)


def cell_style(ws, row, col, value='', bold=False, align='center', size=10, number_format=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(bold=bold, size=size, name='MS Gothic')
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
    c.border = thin_border()
    if number_format:
        c.number_format = number_format
    return c


def set_cell(ws, row, col, value='', bold=False, align='center', size=10, nf=None, fill=None):
    """Set value and style on the TOP-LEFT cell only (safe for merged regions)."""
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(bold=bold, size=size, name='MS Gothic')
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
    c.border = thin_border()
    if nf:
        c.number_format = nf
    if fill:
        c.fill = fill
    return c


def generate_monthly_table(facility_name, desig_no, city_code, year, month,
                            taxable_guests, taxable_amount, tax_amount,
                            daily_data=None,
                            exempt_guests=0, exempt_amount=0):
    wb = Workbook()
    ws = wb.active
    ws.title = "宿泊税月計表"

    hdr_fill = PatternFill(fill_type='solid', fgColor='DDEEFF')

    def txt(row, col, value, align='left', bold=False, size=10):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(bold=bold, size=size, name='MS Gothic')
        c.alignment = Alignment(horizontal=align, vertical='center')

    # ── ヘッダー (行1〜4) ──
    ws.merge_cells('A1:F1')
    txt(1, 1, '宿泊税月計表', align='center', bold=True, size=16)
    ws.row_dimensions[1].height = 32

    ws.merge_cells('A2:B2')
    txt(2, 1, f'市町村コード：{city_code}')
    ws.merge_cells('C2:D2')
    txt(2, 3, f'指定番号：{desig_no}')
    ws.merge_cells('E2:F2')
    txt(2, 5, f'宿泊施設名：{facility_name}')
    ws.row_dimensions[2].height = 18

    ws.merge_cells('A3:F3')
    txt(3, 1, f'令和{reiwa(year)}年　{month}月分　　（単位：人、円）')
    ws.row_dimensions[3].height = 18

    # ── テーブルヘッダー (行5〜6) ──
    # 先にマージしてから値をセット
    ws.merge_cells('A5:A6')
    set_cell(ws, 5, 1, '日付', bold=True, align='center', fill=hdr_fill)

    ws.merge_cells('B5:C5')
    set_cell(ws, 5, 2, '課税対象', bold=True, align='center', fill=hdr_fill)

    ws.merge_cells('D5:E5')
    set_cell(ws, 5, 4, '課税免除', bold=True, align='center', fill=hdr_fill)

    # 行6の個別カラム
    for col, label in [(2, '宿泊者数'), (3, '宿泊料金'), (4, '宿泊者数'), (5, '宿泊料金')]:
        set_cell(ws, 6, col, label, bold=True, align='center', fill=hdr_fill)
    # A6は A5:A6 マージ済みなのでボーダーだけ
    ws.cell(row=6, column=1).border = thin_border()

    ws.row_dimensions[5].height = 20
    ws.row_dimensions[6].height = 20

    # ── 日付行 (行7〜) ──
    days = calendar.monthrange(year, month)[1]
    for d in range(1, days + 1):
        r = 6 + d
        day_info = (daily_data or {}).get(d, {})
        g = day_info.get('guests') or ''
        b = day_info.get('base') or ''
        set_cell(ws, r, 1, d, align='center')
        set_cell(ws, r, 2, g, align='center', nf='#,##0' if g else None)
        set_cell(ws, r, 3, b, align='center', nf='#,##0' if b else None)
        set_cell(ws, r, 4, '')
        set_cell(ws, r, 5, '')
        ws.row_dimensions[r].height = 16

    # ── 合計行 ──
    tr = 7 + days
    set_cell(ws, tr, 1, '合計', bold=True, align='center')
    set_cell(ws, tr, 2, taxable_guests, bold=True, align='center', nf='#,##0')
    set_cell(ws, tr, 3, taxable_amount,  bold=True, align='center', nf='#,##0')
    set_cell(ws, tr, 4, exempt_guests  if exempt_guests  else '', align='center', nf='#,##0')
    set_cell(ws, tr, 5, exempt_amount  if exempt_amount  else '', align='center', nf='#,##0')
    ws.row_dimensions[tr].height = 22

    # ── 税額行 ──
    xr = tr + 1
    ws.merge_cells(f'A{xr}:B{xr}')
    set_cell(ws, xr, 1, '税額（税率２%）', bold=True, align='center')
    ws.cell(row=xr, column=2).border = thin_border()  # マージ範囲の右端
    set_cell(ws, xr, 3, tax_amount, bold=True, align='center', nf='#,##0')
    set_cell(ws, xr, 4, '')
    set_cell(ws, xr, 5, '')
    ws.row_dimensions[xr].height = 22

    # ── 列幅 ──
    for col_letter, w in zip('ABCDEF', [8, 13, 16, 13, 16, 8]):
        ws.column_dimensions[col_letter].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def generate_declaration(facility_name, facility_address, desig_no,
                          applicant_name, applicant_address,
                          year, month,
                          taxable_guests, taxable_amount, tax_amount,
                          exempt_guests=0, exempt_amount=0):
    wb = Workbook()
    ws = wb.active
    ws.title = "宿泊税納入申告書"

    today = date.today()

    def h(row, col, value='', bold=False, align='center', size=10, merge_end_col=None, border=True, nf=None):
        if merge_end_col:
            ws.merge_cells(start_row=row, start_column=col,
                           end_row=row, end_column=merge_end_col)
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(bold=bold, size=size, name='MS Gothic')
        c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
        if border:
            c.border = thin_border()
        if nf:
            c.number_format = nf
        return c

    # Title
    ws.merge_cells('A1:H1')
    c = ws.cell(row=1, column=1, value='宿泊税納入申告書')
    c.font = Font(bold=True, size=18, name='MS Gothic')
    c.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40

    # 指定番号
    ws.merge_cells('F2:H2')
    ws.cell(row=2, column=6, value=f'指定番号　{desig_no}')
    ws.cell(row=2, column=6).font = Font(name='MS Gothic', size=10)
    ws.cell(row=2, column=6).alignment = Alignment(horizontal='right', vertical='center')

    # 宛先
    ws.merge_cells('A3:H3')
    ws.cell(row=3, column=1, value='北谷町長　様')
    ws.cell(row=3, column=1).font = Font(name='MS Gothic', size=12)
    ws.cell(row=3, column=1).alignment = Alignment(horizontal='left', vertical='center')

    # 日付
    ws.merge_cells('F4:H4')
    ws.cell(row=4, column=6, value=f'令和{reiwa(today.year)}年　{today.month}月　{today.day}日')
    ws.cell(row=4, column=6).font = Font(name='MS Gothic', size=10)
    ws.cell(row=4, column=6).alignment = Alignment(horizontal='right', vertical='center')

    # 申告者情報
    ws.merge_cells('A5:B5')
    ws.cell(row=5, column=1, value='申告者の住所').font = Font(name='MS Gothic', size=9)
    ws.cell(row=5, column=1).alignment = Alignment(horizontal='left', vertical='center')
    ws.merge_cells('C5:H5')
    ws.cell(row=5, column=3, value=applicant_address).font = Font(name='MS Gothic', size=10)
    ws.cell(row=5, column=3).alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[5].height = 18

    ws.merge_cells('A6:B6')
    ws.cell(row=6, column=1, value='（法人名称及び代表者名）').font = Font(name='MS Gothic', size=9)
    ws.cell(row=6, column=1).alignment = Alignment(horizontal='left', vertical='center')
    ws.merge_cells('C6:H6')
    ws.cell(row=6, column=3, value=applicant_name).font = Font(name='MS Gothic', size=10)
    ws.cell(row=6, column=3).alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[6].height = 18

    # 施設情報
    fill_gray = PatternFill(fill_type='solid', fgColor='EEEEEE')
    for r, label, val in [
        (7, '所在地', facility_address),
        (8, '名　称', facility_name),
    ]:
        ws.merge_cells(f'A{r}:B{r}')
        c1 = ws.cell(row=r, column=1, value=f'この申告に係る\n宿泊施設　{label}' if r == 7 else f'　　　　　{label}')
        c1.font = Font(name='MS Gothic', size=9)
        c1.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        c1.fill = fill_gray
        ws.merge_cells(f'C{r}:H{r}')
        c2 = ws.cell(row=r, column=3, value=val)
        c2.font = Font(name='MS Gothic', size=10)
        c2.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[r].height = 20

    # Statement
    ws.merge_cells('A9:H9')
    ws.cell(row=9, column=1,
            value='宿泊税の納入について、北谷町宿泊税条例第12条第1項の規定により申告します。')
    ws.cell(row=9, column=1).font = Font(name='MS Gothic', size=10)
    ws.cell(row=9, column=1).alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[9].height = 20

    # Table header
    fill_hdr = PatternFill(fill_type='solid', fgColor='DDEEFF')
    headers = ['年月', '区分', '宿泊者数（延べ数）', '宿泊料金総額（A）', '税率（B）', '税額（A×B）']
    col_map = [1, 2, 3, 4, 6, 7]
    widths_str = [10, 14, 16, 16, 12, 16]

    r = 10
    ws.merge_cells(f'A{r}:A{r+2}')
    c = ws.cell(row=r, column=1,
                value=f'令和{reiwa(year)}年\n{month}月分')
    c.font = Font(bold=True, size=10, name='MS Gothic')
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    c.border = thin_border()
    c.fill = fill_hdr
    ws.row_dimensions[r].height = 24

    for hi, (col, label) in enumerate(zip(col_map[1:], headers[1:])):
        c = ws.cell(row=r, column=col, value=label)
        c.font = Font(bold=True, size=10, name='MS Gothic')
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = thin_border()
        c.fill = fill_hdr

    # Merge 税率 col
    ws.merge_cells(f'E{r}:E{r}')

    # Data rows
    total_guests_all = taxable_guests + exempt_guests
    rows_data = [
        (r+1, '課税対象', taxable_guests, taxable_amount, '100分の2.0', tax_amount),
        (r+2, '課税免除対象', exempt_guests if exempt_guests else '', exempt_amount if exempt_amount else '', '', ''),
        (r+3, '合　計', total_guests_all, '', '', ''),
    ]

    ws.merge_cells(f'A{r+3}:A{r+3}')  # year-month already merged above

    for dr, label, guests, amount, rate, tax in rows_data:
        ws.row_dimensions[dr].height = 22
        cell_style(ws, dr, 2, label, align='center')
        cell_style(ws, dr, 3, guests if guests != '' else '', number_format='#,##0' if guests != '' else None)
        cell_style(ws, dr, 4, amount if amount != '' else '', number_format='#,##0' if amount != '' else None)
        cell_style(ws, dr, 6, rate)
        cell_style(ws, dr, 7, tax if tax != '' else '', bold=(label == '課税対象'), number_format='#,##0' if tax != '' else None)
        cell_style(ws, dr, 5, '')  # empty col
        ws.cell(row=dr, column=1).border = thin_border()

    # Units
    ws.merge_cells(f'C{r}:C{r}')
    note_units_row = r + len(rows_data) + 1
    ws.merge_cells(f'C{note_units_row}:G{note_units_row}')
    ws.cell(row=note_units_row, column=3, value='（単位：人、円）')
    ws.cell(row=note_units_row, column=3).font = Font(name='MS Gothic', size=9)
    ws.cell(row=note_units_row, column=3).alignment = Alignment(horizontal='right')

    # Notes
    notes = [
        '注１　この申告書は、前月中の宿泊について記載し、毎月末日までに提出してください。',
        '注２　この申告書に、内訳となる宿泊年月日ごとに記載された月計表を添付してください。',
        '注３　税率は、町宿泊税100分の1.2と沖縄県宿泊税100分の0.8を合算した税率となっています。',
    ]
    nr = note_units_row + 1
    for i, note in enumerate(notes):
        ws.merge_cells(f'A{nr+i}:H{nr+i}')
        c = ws.cell(row=nr+i, column=1, value=note)
        c.font = Font(name='MS Gothic', size=9)
        c.alignment = Alignment(horizontal='left', vertical='center')
        ws.row_dimensions[nr+i].height = 16

    # Column widths
    for col_letter, w in zip('ABCDEFGH', [12, 16, 16, 16, 6, 13, 16, 8]):
        ws.column_dimensions[col_letter].width = w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def generate_nonyusho(facility_name, facility_address, desig_no, city_code,
                       applicant_name, applicant_address,
                       year, month, tax_amount):
    wb = Workbook()
    ws = wb.active
    ws.title = "宿泊税納入書"

    nm = month % 12 + 1
    ny = year + (1 if month == 12 else 0)
    due_str = f'令和{reiwa(ny)}年{nm}月{calendar.monthrange(ny, nm)[1]}日'

    PCOLS = 9
    GAP   = 1

    hdr_fill = PatternFill(fill_type='solid', fgColor='CCCCCC')
    ts = Side(style='thin')
    ns = Side(style=None)

    def set_outer_border(row, c1, c2, fill=None):
        for col in range(c1, c2 + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = Border(
                left=ts  if col == c1 else ns,
                right=ts if col == c2 else ns,
                top=ts,
                bottom=ts,
            )
            if fill:
                cell.fill = fill

    def mset(row, c1, c2, value='', bold=False, align='center', size=9, fill=None):
        if c1 != c2:
            ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)
        set_outer_border(row, c1, c2, fill=fill)
        c = ws.cell(row=row, column=c1, value=value)
        c.font = Font(bold=bold, size=size, name='MS Gothic')
        c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
        return c

    def scell(row, col, value='', bold=False, align='center', size=9, fill=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = Font(bold=bold, size=size, name='MS Gothic')
        c.alignment = Alignment(horizontal=align, vertical='center')
        c.border = thin_border()
        if fill:
            c.fill = fill
        return c

    def draw_panel(idx, title):
        s = idx * (PCOLS + GAP) + 1

        def a(o): return s + o

        mset(1, a(0), a(8), title, bold=True, size=11, fill=hdr_fill)

        mset(2, a(0), a(2), f'市町村コード\n{city_code}', size=8, fill=hdr_fill)
        mset(2, a(3), a(8), '加入者名\n北谷町会計管理者', size=8, fill=hdr_fill)

        mset(3, a(0), a(2), f'令和{reiwa(year)}年度', size=9, fill=hdr_fill)
        mset(3, a(3), a(4), f'{month}月', size=9, fill=hdr_fill)
        mset(3, a(5), a(8), '■当期　□修正　□当期外', size=7, fill=hdr_fill)

        mset(4, a(0), a(2), '指定番号', size=8, bold=True, fill=hdr_fill)
        mset(4, a(3), a(8), desig_no or '―', size=9)

        mset(5, a(0), a(1), '', fill=hdr_fill)
        for i, lbl in enumerate(['百万', '十万', '万', '千', '百', '十', '円']):
            scell(5, a(2 + i), lbl, size=7, fill=hdr_fill)

        def digit_row(row, label, amount):
            mset(row, a(0), a(1), label, size=8, bold=True)
            if amount is not None:
                digits = [int(d) for d in str(max(0, int(amount))).zfill(7)]
                for i, d in enumerate(digits):
                    scell(row, a(2 + i), d, size=9, align='right')
            else:
                for i in range(7):
                    scell(row, a(2 + i), '', size=9)

        digit_row(6, '税　額', tax_amount)
        digit_row(7, '追滞金', None)
        digit_row(8, '加算金', None)
        digit_row(9, '合計額', tax_amount)

        mset(10, a(0), a(1), '納期限', size=8, bold=True, fill=hdr_fill)
        mset(10, a(2), a(8), due_str, size=9)

        mset(11, a(0), a(8), '特別徴収義務者', size=8, bold=True, fill=hdr_fill)

        scell(12, a(0), '住所', size=8, fill=hdr_fill)
        mset(12, a(1), a(8), facility_address or applicant_address or '', size=8, align='left')

        scell(13, a(0), '氏名', size=8, fill=hdr_fill)
        mset(13, a(1), a(8), applicant_name or '', size=8, align='left')

        scell(14, a(0), '施設名', size=8, fill=hdr_fill)
        mset(14, a(1), a(8), facility_name or '', size=8, align='left')

    for i, t in enumerate(['宿泊税納入書', '宿泊税納入済通知書', '宿泊税領収証書']):
        draw_panel(i, t)

    for r, h in {1: 28, 2: 28, 3: 20, 4: 18, 5: 16,
                 6: 22, 7: 22, 8: 22, 9: 22,
                 10: 18, 11: 16, 12: 18, 13: 18, 14: 18}.items():
        ws.row_dimensions[r].height = h

    for i in range(3):
        base = i * (PCOLS + GAP)
        s = base + 1
        ws.column_dimensions[get_column_letter(s)].width = 5
        ws.column_dimensions[get_column_letter(s + 1)].width = 5
        for j in range(7):
            ws.column_dimensions[get_column_letter(s + 2 + j)].width = 2.8
        if i < 2:
            ws.column_dimensions[get_column_letter(s + 9)].width = 1

    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────── テーマ定数 ───────────
_LOGO_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" width="64" height="64">
  <defs>
    <linearGradient id="bg" x1="0%" y1="0%" x2="100%" y2="100%"><stop offset="0%" style="stop-color:#1B4F72"/><stop offset="100%" style="stop-color:#0E2F44"/></linearGradient>
    <linearGradient id="rf" x1="0%" y1="0%" x2="100%" y2="0%"><stop offset="0%" style="stop-color:#148F77"/><stop offset="100%" style="stop-color:#1ABC9C"/></linearGradient>
  </defs>
  <rect width="100" height="100" rx="18" fill="url(#bg)"/>
  <rect x="22" y="42" width="56" height="40" fill="#2471A3" rx="3"/>
  <polygon points="50,14 20,42 80,42" fill="url(#rf)"/>
  <rect x="48.5" y="7" width="3" height="9" fill="#F1C40F" rx=".5"/>
  <polygon points="51.5,7 51.5,13 58.5,10" fill="#F1C40F"/>
  <rect x="27" y="47" width="11" height="8" fill="white" opacity=".92" rx="1.5"/>
  <rect x="44" y="47" width="12" height="8" fill="#AED6F1" opacity=".92" rx="1.5"/>
  <rect x="62" y="47" width="11" height="8" fill="white" opacity=".92" rx="1.5"/>
  <rect x="27" y="59" width="11" height="8" fill="#AED6F1" opacity=".55" rx="1.5"/>
  <rect x="62" y="59" width="11" height="8" fill="white" opacity=".55" rx="1.5"/>
  <rect x="43" y="62" width="14" height="20" fill="white" opacity=".95" rx="2"/>
  <circle cx="53" cy="73" r="1.5" fill="#2471A3"/>
  <path d="M0 86 Q10 80 22 86 Q34 92 46 86 Q58 80 70 86 Q82 92 100 86 L100 100 L0 100 Z" fill="#1ABC9C" opacity=".45"/>
  <path d="M0 92 Q14 86 28 92 Q42 98 56 92 Q70 86 84 92 Q92 96 100 92 L100 100 L0 100 Z" fill="#1ABC9C" opacity=".3"/>
  <circle cx="78" cy="26" r="12" fill="#F1C40F"/>
  <text x="78" y="31.5" text-anchor="middle" font-family="Arial" font-size="16" fill="#1B4F72" font-weight="900">¥</text>
</svg>"""

_CSS = """<style>
[data-testid="stAppViewContainer"]{background:#EBF2FA;}
[data-testid="stHeader"]{background:transparent;}
.block-container{padding-top:.8rem!important;}
[data-testid="stSidebar"]{background:linear-gradient(160deg,#1B4F72 0%,#0E2F44 100%)!important;}
[data-testid="stSidebar"] *{color:#D6EAF8!important;}
[data-testid="stSidebar"] input{background:rgba(255,255,255,.12)!important;color:white!important;border:1px solid rgba(255,255,255,.25)!important;border-radius:6px!important;}
[data-testid="stSidebar"] .stMarkdown h2{color:white!important;font-size:15px!important;border-bottom:1px solid rgba(255,255,255,.15);padding-bottom:6px;}
[data-testid="stSidebar"] .stMarkdown h3{color:#1ABC9C!important;font-size:11px!important;text-transform:uppercase;letter-spacing:1.5px;}
[data-testid="stSidebar"] hr{border-color:rgba(255,255,255,.15)!important;}
.stTabs [data-baseweb="tab-list"]{background:white;border-radius:12px;padding:5px;gap:4px;box-shadow:0 2px 10px rgba(27,79,114,.1);margin-bottom:4px;}
.stTabs [data-baseweb="tab"]{border-radius:8px;font-weight:600;font-size:14px;padding:9px 22px;color:#7F8C8D;border:none!important;}
.stTabs [aria-selected="true"]{background:linear-gradient(135deg,#1B4F72,#1A8C78)!important;color:white!important;}
.stTabs [data-baseweb="tab-border"]{display:none;}
[data-testid="stMetric"]{background:white;border-radius:14px;padding:20px 24px!important;box-shadow:0 3px 14px rgba(27,79,114,.08);border-bottom:4px solid #1ABC9C;}
[data-testid="stMetricValue"]{color:#1B4F72!important;font-weight:700!important;font-size:2rem!important;}
[data-testid="stMetricLabel"] p{color:#5D6D7E!important;font-size:11px!important;font-weight:600!important;letter-spacing:.5px;text-transform:uppercase;}
[data-testid="stDownloadButton"]>button{background:linear-gradient(135deg,#1B4F72 0%,#1ABC9C 100%)!important;color:white!important;border:none!important;border-radius:10px!important;font-weight:600!important;padding:13px 20px!important;width:100%!important;box-shadow:0 4px 14px rgba(26,188,156,.3)!important;transition:transform .2s,box-shadow .2s!important;font-size:14px!important;letter-spacing:.2px;}
[data-testid="stDownloadButton"]>button:hover{transform:translateY(-2px)!important;box-shadow:0 7px 20px rgba(26,188,156,.45)!important;}
[data-testid="stFileUploaderDropzone"]{background:white!important;border:2px dashed #1ABC9C!important;border-radius:12px!important;}
[data-testid="stDataFrame"]{border-radius:10px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.05);}
[data-testid="stAlert"]{border-radius:10px!important;}
hr{border-color:#D5E8F5!important;}
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:#EBF2FA;}
::-webkit-scrollbar-thumb{background:#1ABC9C;border-radius:3px;}
</style>"""

_HEADER_HTML = f"""<div style="background:linear-gradient(135deg,#1B4F72 0%,#1A8C78 100%);border-radius:16px;padding:22px 32px;margin-bottom:20px;display:flex;align-items:center;gap:20px;box-shadow:0 6px 24px rgba(27,79,114,.22);">
  <div style="flex-shrink:0">{_LOGO_SVG}</div>
  <div style="flex:1">
    <div style="color:white;font-size:26px;font-weight:700;letter-spacing:-.3px;line-height:1.1;margin-bottom:5px;">宿泊税計算システム</div>
    <div style="color:rgba(255,255,255,.78);font-size:13px;">北谷町宿泊税 &nbsp;｜&nbsp; 月計表・申告書・納入書 自動作成ツール</div>
  </div>
  <div style="display:flex;flex-direction:column;gap:6px;align-items:flex-end;">
    <span style="background:rgba(255,255,255,.18);color:white;padding:4px 14px;border-radius:20px;font-size:11px;font-weight:700;letter-spacing:.5px;white-space:nowrap;">🏝 北谷町対応</span>
    <span style="background:#F1C40F;color:#1B4F72;padding:4px 14px;border-radius:20px;font-size:11px;font-weight:700;white-space:nowrap;">税率 町1.2% + 県0.8% = 2.0%</span>
  </div>
</div>"""


def _sec(icon: str, title: str, sub: str = '', color: str = '#1B4F72') -> str:
    """スタイル付きセクションヘッダーHTMLを返す。"""
    sub_html = f'<div style="font-size:12px;color:#7F8C8D;margin-top:2px;">{sub}</div>' if sub else ''
    return (
        f'<div style="background:white;border-radius:12px;padding:14px 20px;margin:8px 0 14px 0;'
        f'box-shadow:0 2px 8px rgba(27,79,114,.07);border-left:4px solid {color};">'
        f'<div style="display:flex;align-items:center;gap:10px;">'
        f'<span style="font-size:20px;">{icon}</span>'
        f'<div><div style="font-weight:700;font-size:15px;color:{color};">{title}</div>{sub_html}</div>'
        f'</div></div>'
    )


def _dl_card(num: str, title: str, desc: str, note: str, icon: str) -> str:
    """ダウンロードカードのHTMLを返す。"""
    return (
        f'<div style="background:white;border-radius:14px;padding:18px 20px 10px;'
        f'box-shadow:0 3px 14px rgba(27,79,114,.09);margin-bottom:10px;">'
        f'<div style="font-size:22px;margin-bottom:6px;">{icon}</div>'
        f'<div style="font-weight:700;color:#1B4F72;font-size:15px;margin-bottom:4px;">{num} {title}</div>'
        f'<div style="font-size:12px;color:#5D6D7E;margin-bottom:3px;">{desc}</div>'
        f'<div style="font-size:11px;color:#95A5A6;">{note}</div>'
        f'</div>'
    )


# ─────────── UI ───────────
st.markdown(_CSS + _HEADER_HTML, unsafe_allow_html=True)

# ── サイドバー：施設・期間設定 ──
with st.sidebar:
    st.header('⚙️ 施設・期間設定')
    city_code = st.text_input('市町村コード', value='473260', help='北谷町：473260')
    desig_no = st.text_input('指定番号', placeholder='例：001')
    facility_name = st.text_input('宿泊施設名', placeholder='例：○○ホテル')
    facility_address = st.text_input('施設所在地', placeholder='例：沖縄県中頭郡北谷町桑江1-1-1')
    applicant_name = st.text_input('申告者名（法人名・代表者名）')
    applicant_address = st.text_input('申告者住所')

    st.divider()
    st.subheader('📅 申告対象年月')
    col_y, col_m = st.columns(2)
    with col_y:
        target_year = st.number_input('年（西暦）', min_value=2025, max_value=2035, value=2026)
    with col_m:
        target_month = st.number_input('月', min_value=1, max_value=12, value=4)

    st.divider()
    st.subheader('💴 金額設定')
    tax_inclusive = st.checkbox(
        'CSVの金額は消費税込み（÷1.1で税抜換算）',
        value=False,
        help='基本利用金額に消費税10%が含まれている場合はチェック'
    )
    meal_per_person = st.number_input(
        '食事代控除（1人1泊・円）',
        min_value=0, value=0, step=100,
        help='食事付きプランがある場合、1人1泊あたりの食事代を入力（宿泊料金から一律控除）'
    )

    st.divider()
    st.subheader('🎓 課税免除（学校行事等）')
    exempt_guests = st.number_input('課税免除 宿泊者数（延べ数）', min_value=0, value=0)
    exempt_amount = st.number_input('課税免除 宿泊料金合計（円）', min_value=0, value=0)

# ── メインタブ ──
tab_upload, tab_calc, tab_export = st.tabs([
    '📁  データ読み込み',
    '📊  計算・確認',
    '📄  書類出力',
])

# ── タブ1: CSV読み込み ──
with tab_upload:
    st.markdown(_sec('📁', '売上CSVファイルの読み込み',
                     '日次予約CSVをアップロード ─ 宿泊日を自動判定し月別に集計します'), unsafe_allow_html=True)
    uploaded = st.file_uploader(
        '日次予約CSVをアップロード',
        type=['csv'],
        help='ホテル管理システムからエクスポートした日次予約一覧CSV（1行＝1予約×1泊）'
    )

    if uploaded:
        raw = uploaded.read()
        items = parse_daily_csv(raw, int(target_year), int(target_month))
        st.session_state['items'] = items
        st.session_state['raw_csv'] = raw  # 検証用に保持

        if items:
            st.success(f'✅ {len(items)} 件の宿泊レコード（予約×泊）を読み込みました')

            # 日別集計プレビュー（税計算前の生データ）
            raw_daily = {}
            for i in items:
                d = i['day']
                if d not in raw_daily:
                    raw_daily[d] = {'reservations': 0, 'guests': 0, 'amount': 0}
                raw_daily[d]['reservations'] += 1
                raw_daily[d]['guests'] += i['guests']
                raw_daily[d]['amount'] += i['amount']

            days_found = sorted(raw_daily.keys())
            st.info(
                f'対象期間：令和{reiwa(target_year)}年{target_month}月　'
                f'｜　有データ日数：{len(days_found)} 日'
                f'（{target_month}/{min(days_found)} 〜 {target_month}/{max(days_found)}）'
            )

            preview_rows = [
                {
                    '日付': f'{target_month}/{d}',
                    '予約件数': raw_daily[d]['reservations'],
                    '宿泊者数（人）': raw_daily[d]['guests'],
                    '基本利用合計（円）': f"{raw_daily[d]['amount']:,}",
                }
                for d in days_found
            ]
            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)
            st.caption('※ 月をまたぐ滞在は宿泊日で自動分割し、対象月分のみを集計しています。')
        else:
            st.warning(f'令和{reiwa(target_year)}年{target_month}月の宿泊データが見つかりませんでした。')
            st.info('サイドバーの「申告対象年月」が正しいか確認してください。')

# ── タブ2: 計算結果 ──
with tab_calc:
    st.markdown(_sec('📊', '宿泊税 計算結果',
                     '日別の課税標準額・宿泊税額を確認します'), unsafe_allow_html=True)

    if 'items' not in st.session_state or not st.session_state['items']:
        st.info('「データ読み込み」タブでCSVファイルをアップロードしてください。')
    else:
        items = st.session_state['items']

        # 日別集計（予約ごとに税計算→日別合算）
        daily_data = compute_daily_agg(
            items,
            meal_per_person=int(meal_per_person),
            tax_inclusive=tax_inclusive,
        )

        total_taxable_guests = sum(v['guests'] for v in daily_data.values())
        total_taxable_base   = sum(v['base']   for v in daily_data.values())
        total_tax            = sum(v['tax']    for v in daily_data.values())

        st.session_state['calc'] = {
            'total_guests': total_taxable_guests,
            'total_base': total_taxable_base,
            'total_tax': total_tax,
        }
        st.session_state['daily_data'] = daily_data

        # サマリー指標
        c1, c2, c3 = st.columns(3)
        c1.metric('課税対象 宿泊者数（延べ）', f'{total_taxable_guests:,} 人')
        c2.metric('宿泊料金総額（課税標準額）', f'¥{total_taxable_base:,}')
        c3.metric('🏷️ 宿泊税額', f'¥{total_tax:,}')

        st.divider()

        # 日別明細テーブル
        st.markdown(_sec('📅', '日別 計算明細', '', '#1A8C78'), unsafe_allow_html=True)
        detail_rows = [
            {
                '日付': f'{target_month}/{d}',
                '宿泊者数（人）': daily_data[d]['guests'],
                '課税標準額合計（円）': f"{daily_data[d]['base']:,}",
                '宿泊税（円）': f"{daily_data[d]['tax']:,}",
            }
            for d in sorted(daily_data.keys())
        ]
        st.dataframe(pd.DataFrame(detail_rows), use_container_width=True, hide_index=True)
        st.caption('※ 課税標準額は1人当たり宿泊料金を千円未満切り捨てた額×人数の合計です。')

        st.divider()

        # 申告書サマリー
        st.markdown(_sec('📝', '申告書サマリー', '課税対象・課税免除・合計の区分ごとの集計', '#E67E22'), unsafe_allow_html=True)
        total_all = total_taxable_guests + exempt_guests
        summary_df = pd.DataFrame([
            {'区分': '課税対象', '宿泊者数（延べ数）': f'{total_taxable_guests:,} 人',
             '宿泊料金総額（A）': f'¥{total_taxable_base:,}', '税率（B）': '100分の2.0',
             '税額（A×B）': f'¥{total_tax:,}'},
            {'区分': '課税免除対象', '宿泊者数（延べ数）': f'{exempt_guests:,} 人',
             '宿泊料金総額（A）': f'¥{exempt_amount:,}' if exempt_amount else '—',
             '税率（B）': '—', '税額（A×B）': '—'},
            {'区分': '合　計', '宿泊者数（延べ数）': f'{total_all:,} 人',
             '宿泊料金総額（A）': '—', '税率（B）': '—', '税額（A×B）': '—'},
        ])
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        # ── 検証セクション ──
        st.divider()
        st.markdown(_sec('🔍', '計算検証', 'CSV小計との照合 ＆ 予約別の計算ステップを確認', '#8E44AD'), unsafe_allow_html=True)

        with st.expander('検証データを開く（クリックで展開）', expanded=False):
            if 'raw_csv' not in st.session_state:
                st.info('「データ読み込み」タブでCSVを再アップロードすると検証データが表示されます。')
            else:
                v_detail, v_sub = parse_for_verification(
                    st.session_state['raw_csv'],
                    int(target_year), int(target_month),
                    meal_per_person=int(meal_per_person),
                    tax_inclusive=tax_inclusive,
                )

                # ── 日別 照合テーブル ──
                st.subheader('① 日別集計の照合（当システム vs CSV小計）')
                st.caption('人数が一致していれば、宿泊日の計算（月またぎ含む）が正しいことを意味します。')
                comp_rows = []
                all_days = sorted(set(list(daily_data.keys()) + list(v_sub.keys())))
                for d in all_days:
                    our_g = daily_data.get(d, {}).get('guests', 0)
                    our_b = daily_data.get(d, {}).get('base', 0)
                    our_t = daily_data.get(d, {}).get('tax', 0)
                    csv_g = v_sub.get(d, {}).get('guests', '―')
                    match = '✅ 一致' if our_g == csv_g else f'❌ 不一致（CSV:{csv_g}）'
                    comp_rows.append({
                        '日付': f'{target_month}/{d}',
                        '当システム 人数': our_g,
                        'CSV 小計 人数': csv_g,
                        '照合': match,
                        '課税標準額合計（円）': f'{our_b:,}',
                        '宿泊税（円）': f'{our_t:,}',
                    })
                st.dataframe(pd.DataFrame(comp_rows), use_container_width=True, hide_index=True)

                # 照合サマリー
                mismatches = [r for r in comp_rows if '❌' in r['照合']]
                if mismatches:
                    st.error(f'❌ {len(mismatches)} 日で不一致があります。CSVの日付や申告月の設定を確認してください。')
                else:
                    st.success(f'✅ 全 {len(comp_rows)} 日で人数が一致しました。日付計算は正確です。')

                st.divider()

                # ── 予約別 計算ステップ ──
                st.subheader('② 予約別 計算ステップ明細')
                st.caption(
                    '**課税標準額/人** = 基本利用額 ÷ 人数 を千円未満切捨 　'
                    '**宿泊税** = 課税標準額/人 × 2% × 人数（上限2,000円/人）'
                )
                if v_detail:
                    df_v = pd.DataFrame(v_detail)
                    # 数値列にカンマ書式
                    for col in ['基本利用（円）', '1人当たり料金', '課税標準額/人（千円切捨）', '課税標準額合計', '宿泊税（円）']:
                        df_v[col] = df_v[col].apply(lambda x: f'{x:,}')
                    st.dataframe(df_v, use_container_width=True, hide_index=True)

                    # 検証Excelダウンロード
                    def _make_verify_excel(detail, comparison):
                        wb = Workbook()
                        # Sheet1: 予約別明細
                        ws1 = wb.active
                        ws1.title = '予約別計算明細'
                        cols1 = list(detail[0].keys())
                        for ci, h in enumerate(cols1, 1):
                            c = ws1.cell(row=1, column=ci, value=h)
                            c.font = Font(bold=True, size=9, name='MS Gothic')
                            c.fill = PatternFill(fill_type='solid', fgColor='DDEEFF')
                            c.border = thin_border()
                            c.alignment = Alignment(horizontal='center', vertical='center')
                        for ri, row in enumerate(detail, 2):
                            for ci, key in enumerate(cols1, 1):
                                c = ws1.cell(row=ri, column=ci, value=row[key])
                                c.font = Font(size=9, name='MS Gothic')
                                c.border = thin_border()
                                c.alignment = Alignment(horizontal='center', vertical='center')
                        for col in ws1.columns:
                            ws1.column_dimensions[col[0].column_letter].width = 16
                        # Sheet2: 日別照合
                        ws2 = wb.create_sheet('日別照合')
                        cols2 = list(comparison[0].keys())
                        for ci, h in enumerate(cols2, 1):
                            c = ws2.cell(row=1, column=ci, value=h)
                            c.font = Font(bold=True, size=9, name='MS Gothic')
                            c.fill = PatternFill(fill_type='solid', fgColor='DDEEFF')
                            c.border = thin_border()
                            c.alignment = Alignment(horizontal='center', vertical='center')
                        for ri, row in enumerate(comparison, 2):
                            for ci, key in enumerate(cols2, 1):
                                c = ws2.cell(row=ri, column=ci, value=row[key])
                                c.font = Font(size=9, name='MS Gothic')
                                c.border = thin_border()
                                c.alignment = Alignment(horizontal='center', vertical='center')
                        for col in ws2.columns:
                            ws2.column_dimensions[col[0].column_letter].width = 18
                        buf = io.BytesIO()
                        wb.save(buf)
                        buf.seek(0)
                        return buf

                    vbuf = _make_verify_excel(v_detail, comp_rows)
                    st.download_button(
                        label='📥 検証Excelをダウンロード（予約別明細＋日別照合）',
                        data=vbuf,
                        file_name=f'宿泊税検証_{target_year}{target_month:02d}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    )
                else:
                    st.info('対象月の予約データがありません。')

# ── タブ3: 書類出力 ──
with tab_export:
    st.markdown(_sec('📄', '書類出力', '月計表・申告書・納入書を Excel でダウンロードします'), unsafe_allow_html=True)

    if 'calc' not in st.session_state:
        st.info('「計算結果」タブで計算を確認してから書類を出力してください。')
    else:
        calc = st.session_state['calc']
        reiwa_str = f'令和{reiwa(target_year)}年{target_month}月分'

        st.info(f'対象期間：{reiwa_str}　｜　宿泊税額：¥{calc["total_tax"]:,}')

        col_a, col_b, col_c = st.columns(3)

        with col_a:
            st.markdown(_dl_card('①', '宿泊税月計表',
                                 '申告書に添付する月別日次計表',
                                 '日付ごとの宿泊者数・料金が自動入力', '📋'),
                        unsafe_allow_html=True)
            monthly_buf = generate_monthly_table(
                facility_name=facility_name or '（未入力）',
                desig_no=desig_no or '―',
                city_code=city_code or '―',
                year=target_year,
                month=target_month,
                taxable_guests=calc['total_guests'],
                taxable_amount=calc['total_base'],
                tax_amount=calc['total_tax'],
                daily_data=st.session_state.get('daily_data'),
                exempt_guests=exempt_guests,
                exempt_amount=exempt_amount,
            )
            st.download_button(
                label='📥 月計表をダウンロード（Excel）',
                data=monthly_buf,
                file_name=f'宿泊税月計表_{target_year}{target_month:02d}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
            )

        with col_b:
            st.markdown(_dl_card('②', '宿泊税納入申告書',
                                 '北谷町税務課に提出する申告書',
                                 '提出期限：翌月末日（土日祝の場合は翌平日）', '📑'),
                        unsafe_allow_html=True)
            decl_buf = generate_declaration(
                facility_name=facility_name or '（未入力）',
                facility_address=facility_address or '',
                desig_no=desig_no or '―',
                applicant_name=applicant_name or '',
                applicant_address=applicant_address or '',
                year=target_year,
                month=target_month,
                taxable_guests=calc['total_guests'],
                taxable_amount=calc['total_base'],
                tax_amount=calc['total_tax'],
                exempt_guests=exempt_guests,
                exempt_amount=exempt_amount,
            )
            st.download_button(
                label='📥 申告書をダウンロード（Excel）',
                data=decl_buf,
                file_name=f'宿泊税納入申告書_{target_year}{target_month:02d}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
            )

        with col_c:
            st.markdown(_dl_card('③', '宿泊税納入書',
                                 '金融機関での納付に使う3枚綴り',
                                 '納入書・通知書・領収証書の3面構成', '🏦'),
                        unsafe_allow_html=True)
            nonyusho_buf = generate_nonyusho(
                facility_name=facility_name or '（未入力）',
                facility_address=facility_address or '',
                desig_no=desig_no or '―',
                city_code=city_code or '473260',
                applicant_name=applicant_name or '',
                applicant_address=applicant_address or '',
                year=target_year,
                month=target_month,
                tax_amount=calc['total_tax'],
            )
            st.download_button(
                label='📥 納入書をダウンロード（Excel）',
                data=nonyusho_buf,
                file_name=f'宿泊税納入書_{target_year}{target_month:02d}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
            )

        st.divider()
        st.markdown("""
<div style="background:white;border-radius:12px;padding:18px 24px;box-shadow:0 2px 8px rgba(27,79,114,.07);">
<div style="font-weight:700;color:#1B4F72;margin-bottom:10px;font-size:14px;">📮 提出方法</div>
<ul style="margin:0;padding-left:20px;color:#5D6D7E;font-size:13px;line-height:1.9;">
<li>月計表と申告書をセットで北谷町税務課へ提出</li>
<li>提出先：窓口持参 ／ 郵送 ／ eLTAX 利用可</li>
<li>電子申告は eLTAX（PCdeskNext）から申告可能</li>
</ul>
</div>
""", unsafe_allow_html=True)
