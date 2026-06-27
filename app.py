import streamlit as st
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import io
import csv
import math
import calendar
import os
import json
from datetime import date, timedelta, datetime as _dt
import re

st.set_page_config(page_title="宿泊税スマート", page_icon="🏨", layout="wide")

TAX_RATE = 0.02
MAX_TAX_PER_PERSON = 2000
AGODA_COMMISSION             = 0.12   # Agoda通常手数料（Private Sale等）
AGODA_COMMISSION_INTELLIGENCE = 0.09  # Agoda Intelligence チャネルの手数料
AIRBNB_COMMISSION            = 0.155  # Airbnb手数料（EarningsはこれをCSV記載額から控除済み）


def _agoda_commission(note: str) -> float:
    """V列(備考1)の内容からAgoda手数料率を返す。Agoda Intelligenceなら9%、それ以外は12%。"""
    if 'Agoda Intelligence' in note:
        return AGODA_COMMISSION_INTELLIGENCE
    return AGODA_COMMISSION

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_DECLARATION_TPL  = os.path.join(_BASE_DIR, '納入申告書テンプレ.xlsx')
_NONYUSHO_TPL     = os.path.join(_BASE_DIR, '宿泊税納入書テンプレ.xlsx')
_RECEIPT_TPL_A    = os.path.join(_BASE_DIR, '領収書タイプA税込.xlsx')
_RECEIPT_TPL_B    = os.path.join(_BASE_DIR, '領収書タイプB税別途.xlsx')
_RECEIPT_TPL_C    = os.path.join(_BASE_DIR, '領収書タイプC税全コミ.xlsx')
_RECEIPT_TPL_D    = os.path.join(_BASE_DIR, '領収書タイプD宿泊税（税のみ）.xlsx')
_MONTHLY_TPL      = os.path.join(_BASE_DIR, '宿泊税月計表テンプレ.xlsx')
_SETTINGS_FILE    = os.path.join(_BASE_DIR, 'settings.json')


def load_settings():
    if os.path.exists(_SETTINGS_FILE):
        try:
            with open(_SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_settings(data: dict):
    with open(_SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def reiwa(year):
    return year - 2018


def three_month_period(last_year, last_month):
    """申告最終月から遡った3ヶ月分の (year, month) リストを返す（昇順）"""
    result = []
    for offset in (2, 1, 0):
        m = last_month - offset
        y = last_year
        if m <= 0:
            m += 12
            y -= 1
        result.append((y, m))
    return result


def _parse_hotel_date(s: str):
    """チェックイン日付を MM/DD/YYYY と YYYY/M/D の両形式に対応してパースする。"""
    s = s.strip()
    for fmt in ('%m/%d/%Y', '%Y/%m/%d'):
        try:
            return _dt.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError(f'日付フォーマット不明: {s!r}')


def parse_hotel_csv(file_bytes, target_year, target_month):
    """ホテルシステムCSV（43列・MM/DD/YYYY または YYYY/M/D 形式）を解析し、対象月の宿泊レコードを返す。
    同一(予約ID, 泊目)の大人行・子供行を集約。金額ゼロ時は料金合計額÷泊数でフォールバック。
    Returns: list of {day, guests, amount}
    """
    content = file_bytes.decode('cp932', errors='replace')

    grouped = {}  # (予約ID, 泊目) -> {day, guests, amount, site, total_fee, total_nights}

    reader = csv.reader(io.StringIO(content))
    next(reader, None)  # ヘッダーをスキップ
    for parts in reader:
        if len(parts) < 38 or not parts[0]:
            continue
        try:
            night_num    = int(parts[3])
            total_nights = int(parts[7])
            ci           = _parse_hotel_date(parts[4])
            night_date   = ci + timedelta(days=night_num - 1)
        except (ValueError, IndexError):
            continue

        if night_date.year != target_year or night_date.month != target_month:
            continue

        try:
            guests = int(parts[18]) + int(parts[19])  # 大人 + 子供（幼児は課税対象外）
        except (ValueError, IndexError):
            continue
        if guests <= 0:
            continue

        try:
            amount = int(parts[36] or 0) + int(parts[37] or 0)  # 大人合計額 + 子供合計額
        except (ValueError, IndexError):
            amount = 0

        try:
            total_fee = int(str(parts[25]).replace(',', '') or 0)
        except (ValueError, IndexError):
            total_fee = 0

        site = parts[8] if len(parts) > 8 else ''
        note = parts[21] if len(parts) > 21 else ''
        key  = (parts[0], night_num)

        if key not in grouped:
            grouped[key] = {
                'day': night_date.day, 'guests': guests, 'amount': 0,
                'site': site, 'note': note, 'total_fee': total_fee, 'total_nights': total_nights,
            }
        grouped[key]['amount'] += amount

    result = []
    for data in grouped.values():
        amount = data['amount']
        if amount == 0 and data['total_fee'] > 0 and data['total_nights'] > 0:
            amount = round(data['total_fee'] / data['total_nights'])
        if data['site'].lower() == 'agoda':
            amount = round(amount / (1 - _agoda_commission(data.get('note', ''))))
        result.append({'day': data['day'], 'guests': data['guests'], 'amount': amount})

    return result


def detect_csv_format(file_bytes):
    """CSVの1行目を見てフォーマットを判定: 'airbnb' / 'hotel_total' / 'hotel'"""
    try:
        first_line = file_bytes.decode('utf-8-sig', errors='replace').splitlines()[0]
        if 'Confirmation code' in first_line or 'Confirmation Code' in first_line:
            return 'airbnb'
    except Exception:
        pass
    # 75列以上 = ReservationTotalList形式（1行=1予約）
    try:
        lines = file_bytes.decode('cp932', errors='replace').splitlines()
        for line in lines[1:4]:
            if len(line.split(',')) > 60:
                return 'hotel_total'
    except Exception:
        pass
    return 'hotel'


def parse_hotel_total_csv(file_bytes, target_year, target_month):
    """ReservationTotalList形式（1行=1予約、75列）を解析し、対象月の宿泊レコードを返す。
    各予約を宿泊日ごとに展開し、合計額÷泊数で1泊あたりの金額を算出する。
    Returns: list of {day, guests, amount}
    """
    content = file_bytes.decode('cp932', errors='replace')
    result = []

    reader = csv.reader(io.StringIO(content))
    next(reader, None)  # ヘッダーをスキップ
    for parts in reader:
        if len(parts) < 38 or not parts[0] or not parts[4]:
            continue
        try:
            total_nights = int(parts[7])
            ci           = _parse_hotel_date(parts[4])
            guests       = int(parts[18] or 0) + int(parts[19] or 0)
        except (ValueError, IndexError):
            continue
        if guests <= 0 or total_nights <= 0:
            continue

        try:
            total_amt = int(parts[36] or 0) + int(parts[37] or 0)
        except (ValueError, IndexError):
            total_amt = 0
        if total_amt == 0:
            try:
                total_amt = int(str(parts[25]).replace(',', '') or 0)
            except (ValueError, IndexError):
                pass

        per_night = round(total_amt / total_nights)
        site = parts[8] if len(parts) > 8 else ''
        note = parts[21] if len(parts) > 21 else ''
        if site.lower() == 'agoda':
            per_night = round(per_night / (1 - _agoda_commission(note)))

        for night in range(total_nights):
            night_date = ci + timedelta(days=night)
            if night_date.year == target_year and night_date.month == target_month:
                result.append({'day': night_date.day, 'guests': guests, 'amount': per_night})

    return result


def parse_airbnb_csv(file_bytes, target_year, target_month):
    """Airbnb予約CSVを解析し、対象月の宿泊レコードを返す。
    1予約行を宿泊日ごとに展開し、Earnings÷泊数で1泊あたりの金額を算出する。
    Returns: list of {day, guests, amount}
    """
    for enc in ['utf-8-sig', 'utf-8', 'shift_jis']:
        try:
            content = file_bytes.decode(enc)
            break
        except Exception:
            continue
    else:
        return []

    try:
        df = pd.read_csv(io.StringIO(content))
    except Exception:
        return []

    # 列名の前後スペースを除去
    df.columns = df.columns.str.strip()

    items = []
    for _, row in df.iterrows():
        # キャンセル済みは除外
        status = str(row.get('Status', ''))
        if 'cancel' in status.lower():
            continue

        # 人数（大人＋子供、幼児は宿泊税対象外のため除く）
        try:
            adults   = int(row.get('# of adults',   0) or 0)
            children = int(row.get('# of children', 0) or 0)
            guests   = adults + children
        except (ValueError, TypeError):
            guests = 1
        if guests <= 0:
            guests = 1

        # 泊数
        try:
            nights = int(row.get('# of nights', 0) or 0)
        except (ValueError, TypeError):
            continue
        if nights <= 0:
            continue

        # チェックイン日 (DD/MM/YYYY または D/M/YYYY)
        start_str = str(row.get('Start date', '')).strip()
        try:
            checkin = _dt.strptime(start_str, '%d/%m/%Y').date()
        except ValueError:
            continue

        # Earnings（手数料控除後）→ グロスアップして素泊まり料金を算出
        earnings_str = re.sub(r'[¥\s,]', '', str(row.get('Earnings', '0')))
        try:
            total_earnings = int(float(earnings_str))
        except (ValueError, TypeError):
            continue
        gross_total = round(total_earnings / (1 - AIRBNB_COMMISSION))
        per_night = gross_total // nights

        # 各宿泊日ごとにレコードを生成（対象月のみ）
        for n in range(nights):
            night_date = checkin + timedelta(days=n)
            if night_date.year == target_year and night_date.month == target_month:
                items.append({
                    'day': night_date.day,
                    'guests': guests,
                    'amount': per_night,
                })

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


def parse_airbnb_csv_for_verification(file_bytes, target_year, target_month,
                                       meal_per_person=0, tax_inclusive=False):
    """Airbnb CSVを検証用に解析し、予約別明細と月またぎ除外データを返す。
    Returns: (detail_rows, excluded_rows)
    """
    for enc in ['utf-8-sig', 'utf-8', 'shift_jis']:
        try:
            content = file_bytes.decode(enc)
            break
        except Exception:
            continue
    else:
        return [], []

    try:
        df = pd.read_csv(io.StringIO(content))
    except Exception:
        return [], []

    df.columns = df.columns.str.strip()

    detail_rows = []
    excluded_rows = []

    for _, row in df.iterrows():
        status = str(row.get('Status', ''))
        if 'cancel' in status.lower():
            continue

        conf_code = str(row.get('Confirmation code', '')).strip()

        try:
            adults   = int(row.get('# of adults',   0) or 0)
            children = int(row.get('# of children', 0) or 0)
            guests   = adults + children
        except (ValueError, TypeError):
            guests = 1
        if guests <= 0:
            guests = 1

        try:
            nights = int(row.get('# of nights', 0) or 0)
        except (ValueError, TypeError):
            continue
        if nights <= 0:
            continue

        start_str = str(row.get('Start date', '')).strip()
        try:
            checkin = _dt.strptime(start_str, '%d/%m/%Y').date()
        except ValueError:
            continue

        checkout = checkin + timedelta(days=nights)

        earnings_str = re.sub(r'[¥\s,]', '', str(row.get('Earnings', '0')))
        try:
            total_earnings = int(float(earnings_str))
        except (ValueError, TypeError):
            continue
        gross_total = round(total_earnings / (1 - AIRBNB_COMMISSION))
        per_night   = gross_total // nights

        for n in range(nights):
            night_date = checkin + timedelta(days=n)
            if night_date.year == target_year and night_date.month == target_month:
                base_pp, _, tax = calc_tax(guests, per_night,
                                           meal_per_person=meal_per_person,
                                           tax_inclusive=tax_inclusive)
                detail_rows.append({
                    '予約コード':                conf_code,
                    '宿泊日':                   night_date.strftime('%m/%d'),
                    'チェックイン':              checkin.strftime('%m/%d'),
                    'チェックアウト':            checkout.strftime('%m/%d'),
                    '人数（人）':                guests,
                    'Earnings（CSV記載額）':     total_earnings,
                    'グロスアップ後（÷0.845）': gross_total,
                    '1泊料金（円）':             per_night,
                    '課税標準額/人（千円切捨）': base_pp,
                    '宿泊税（円）':              tax,
                })
            else:
                excluded_rows.append({
                    '予約コード':      conf_code,
                    '宿泊日（除外）':  night_date.strftime('%Y/%m/%d'),
                    'チェックイン':    checkin.strftime('%Y/%m/%d'),
                    'チェックアウト':  checkout.strftime('%Y/%m/%d'),
                    '人数（人）':      guests,
                    '除外理由':        f'{night_date.month}月のため対象外',
                })

    return detail_rows, excluded_rows


def compute_daily_agg(items, meal_per_person=0, tax_inclusive=False):
    """予約ごとに宿泊税を計算し、日別に集計して返す。
    Returns: {day: {guests, base, tax, amount}}
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
            daily[d] = {'guests': 0, 'base': 0, 'tax': 0, 'amount': 0}
        daily[d]['guests'] += item['guests']
        daily[d]['base']   += base
        daily[d]['tax']    += tax
        daily[d]['amount'] += item['amount']
    return daily


def recompute_daily_from_edit(edited_df, month):
    """data_editor の出力 DataFrame から日別集計を再構築する。
    1日単位の合計人数・合計宿泊料金を使って宿泊税を計算する。
    Returns: {day: {guests, base, tax, amount}}
    """
    result = {}
    for _, row in edited_df.iterrows():
        try:
            day = int(str(row['日付']).split('/')[-1])
        except (ValueError, TypeError):
            continue
        guests = max(0, int(row.get('宿泊者数', 0) or 0))
        amount = max(0, int(row.get('宿泊料金合計（円）', 0) or 0))
        if guests > 0 and amount > 0:
            per_person = amount / guests
            base_per   = math.floor(per_person / 1000) * 1000
            tax_pp     = min(base_per * TAX_RATE, MAX_TAX_PER_PERSON)
            base       = int(base_per * guests)
            tax        = int(tax_pp  * guests)
        else:
            base = tax = 0
        result[day] = {'guests': guests, 'base': base, 'tax': tax, 'amount': amount}
    return result


def parse_hotel_csv_for_verification(file_bytes, target_year, target_month,
                                      meal_per_person=0, tax_inclusive=False):
    """ホテルシステムCSVの検証用解析。(予約ID, 泊目)単位で集約し、
    日別集計・予約別計算ステップ・除外データ・Agoda補正明細を返す。
    Returns:
        detail_rows    : list[dict] ─ 予約-泊ごとの計算ステップ
        daily_subtotals: {day: {guests, amount}} ─ 日別集計
        excluded_rows  : list[dict] ─ 対象月外として除外した宿泊夜
        agoda_rows     : list[dict] ─ Agoda補正を適用した予約の明細
    """
    content = file_bytes.decode('cp932', errors='replace')
    grouped      = {}
    excluded_map = {}  # (parts[0], night_num) -> 除外レコード情報

    reader = csv.reader(io.StringIO(content))
    next(reader, None)
    for parts in reader:
        if len(parts) < 38 or not parts[0]:
            continue
        try:
            night_num    = int(parts[3])
            total_nights = int(parts[7])
            ci           = _parse_hotel_date(parts[4])
            night_date   = ci + timedelta(days=night_num - 1)
        except (ValueError, IndexError):
            continue

        key = (parts[0], night_num)

        if night_date.year != target_year or night_date.month != target_month:
            # 対象月外 → 除外リストに1回だけ記録
            if key not in excluded_map:
                co_date = ci + timedelta(days=total_nights)
                excluded_map[key] = {
                    'res_no':       parts[2],
                    'site':         parts[8] or '直接',
                    'guest_name':   parts[12],
                    'ci_date':      ci,
                    'co_date':      co_date,
                    'night_date':   night_date,
                    'night_num':    night_num,
                    'total_nights': total_nights,
                }
            continue

        try:
            guests = int(parts[18]) + int(parts[19])
        except (ValueError, IndexError):
            continue
        if guests <= 0:
            continue
        try:
            amount = int(parts[36] or 0) + int(parts[37] or 0)
        except (ValueError, IndexError):
            amount = 0
        try:
            total_fee = int(str(parts[25]).replace(',', '') or 0)
        except (ValueError, IndexError):
            total_fee = 0

        if key not in grouped:
            grouped[key] = {
                'day':          night_date.day,
                'date_str':     f'{target_month}/{night_date.day}',
                'res_no':       parts[2],
                'site':         parts[8] or '直接',
                'guest_name':   parts[12],
                'guests':       guests,
                'amount':       0,
                'note':         parts[21] if len(parts) > 21 else '',
                'total_fee':    total_fee,
                'total_nights': total_nights,
            }
        grouped[key]['amount'] += amount

    detail_rows     = []
    daily_subtotals = {}
    agoda_rows      = []

    for data in grouped.values():
        amount_raw = data['amount']
        if amount_raw == 0 and data['total_fee'] > 0 and data['total_nights'] > 0:
            amount_raw = round(data['total_fee'] / data['total_nights'])

        agoda_adj = data['site'].lower() == 'agoda'
        if agoda_adj:
            amount_gross = round(amount_raw / (1 - _agoda_commission(data.get('note', ''))))
            agoda_rows.append({
                '宿泊日':                   data['date_str'],
                '予約番号':                 data['res_no'],
                '宿泊者名':                 data['guest_name'],
                '人数':                     data['guests'],
                'CSV記載額（控除後）':      amount_raw,
                'グロスアップ後（円）':      amount_gross,
                '補正額（手数料分）':        amount_gross - amount_raw,
            })
            amount = amount_gross
        else:
            amount = amount_raw

        adj_amount   = (amount / 1.10) if tax_inclusive else amount
        net          = adj_amount - meal_per_person * data['guests']
        per_person   = net / data['guests']
        base_per     = math.floor(per_person / 1000) * 1000
        tax_pp       = min(base_per * TAX_RATE, MAX_TAX_PER_PERSON)
        tax_total    = int(tax_pp * data['guests'])

        detail_rows.append({
            '宿泊日':                data['date_str'],
            '予約番号':              data['res_no'],
            '予約サイト':            data['site'],
            '宿泊者名':              data['guest_name'],
            '人数':                  data['guests'],
            '1泊料金（円）':         amount_raw,
            'Agoda補正後（円）':     amount if agoda_adj else '',
            '1人当たり料金':         int(per_person),
            '課税標準額/人（千円切捨）': int(base_per),
            '課税標準額合計':        int(base_per * data['guests']),
            '宿泊税（円）':          tax_total,
        })

        d = data['day']
        if d not in daily_subtotals:
            daily_subtotals[d] = {'guests': 0, 'amount': 0}
        daily_subtotals[d]['guests'] += data['guests']
        daily_subtotals[d]['amount'] += amount

    detail_rows.sort(key=lambda r: (int(r['宿泊日'].split('/')[1]), r['予約番号']))
    agoda_rows.sort(key=lambda r: (r['宿泊日'], r['予約番号']))

    # 除外データ（宿泊日順）
    excluded_rows = [
        {
            '予約番号':          d['res_no'],
            '予約サイト':        d['site'],
            '宿泊者名':          d['guest_name'],
            'チェックイン':      d['ci_date'].strftime('%m/%d'),
            'チェックアウト':    d['co_date'].strftime('%m/%d'),
            '総泊数':            d['total_nights'],
            '泊目':              d['night_num'],
            '除外された宿泊日':  d['night_date'].strftime('%m/%d（%m月）'),
            '除外理由':          f"対象月（{target_month}月）外の宿泊夜",
        }
        for d in sorted(excluded_map.values(), key=lambda x: (x['night_date'], x['res_no']))
    ]

    return detail_rows, daily_subtotals, excluded_rows, agoda_rows


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
                            daily_data=None, exempt_guests=0, exempt_amount=0,
                            _workbook=None):
    """月計表シートを生成。_workbook 指定時はそのブックにシートを追加してブックを返す。"""
    if _workbook is not None:
        wb = _workbook
        ws = wb.create_sheet(f'{reiwa(year)}年{month}月')
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = f'{reiwa(year)}年{month}月'

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

    if _workbook is not None:
        return wb  # 呼び出し元がsaveする

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def generate_monthly_from_template(facility_name, desig_no, months_calc_list):
    """宿泊税月計表テンプレ.xlsx を使い月ごとにシートを複製して返す。"""
    wb = load_workbook(_MONTHLY_TPL)
    tpl_ws = wb.active
    tpl_name = tpl_ws.title

    for mc in months_calc_list:
        year       = mc['year']
        month      = mc['month']
        daily_data = mc.get('daily_data', {})

        new_ws       = wb.copy_worksheet(tpl_ws)
        new_ws.title = f'令和{reiwa(year)}年{month}月'

        new_ws['B1'] = desig_no or ''
        new_ws['B2'] = facility_name or ''
        new_ws['A3'] = f'令和{reiwa(year)}年'
        new_ws['C3'] = f'{month}月分'

        for day in range(1, 32):
            row     = 6 + day
            info    = daily_data.get(day, {})
            guests  = info.get('guests')
            amount  = info.get('base')
            if guests:
                new_ws.cell(row=row, column=2).value = guests
            if amount:
                new_ws.cell(row=row, column=3).value = amount

    del wb[tpl_name]

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


def generate_declaration_from_template(facility_name, facility_address, desig_no,
                                       applicant_name, applicant_address,
                                       months_data):
    """納入申告書テンプレ.xlsx を読み込み、最大3ヶ月分のデータを書き込んで返す。
    months_data: list of dict (1〜3件、時系列順)
      各 dict: {year, month, taxable_guests, taxable_amount, tax_amount,
                exempt_guests=0, exempt_amount=0}
    """
    wb = load_workbook(_DECLARATION_TPL)
    ws = wb.active
    today = date.today()

    def w(row, col, val):
        ws.cell(row=row, column=col).value = val

    # ヘッダー部（共通）
    w(3, 7, desig_no or '')
    w(5, 5, f'令和{reiwa(today.year)}年　{today.month}月　{today.day}日')
    w(9, 2, applicant_address or '')
    w(9, 5, applicant_name or '')
    w(14, 4, facility_address or '')
    w(15, 4, facility_name or '')

    # 3ブロック: 行開始 19, 24, 29
    for block_row, md in zip([19, 24, 29], months_data):
        year            = md['year']
        month           = md['month']
        taxable_guests  = md.get('taxable_guests', 0)
        taxable_amount  = md.get('taxable_amount', 0)
        tax_amount      = md.get('tax_amount', 0)
        exempt_guests   = md.get('exempt_guests', 0)
        exempt_amount   = md.get('exempt_amount', 0)

        w(block_row,   1, f'令和{reiwa(year)}年')
        w(block_row+2, 1, f'{month}月分')
        w(block_row,   3, f'{taxable_guests:,}人')
        w(block_row,   4, f'{taxable_amount:,}円')
        w(block_row,   7, f'{tax_amount:,}円')
        if exempt_guests:
            w(block_row+1, 3, f'{exempt_guests:,}人')
        if exempt_amount:
            w(block_row+1, 4, f'{exempt_amount:,}円')
        w(block_row+2, 3, f'{taxable_guests + exempt_guests:,}人')

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def generate_nonyusho_from_template(facility_name, facility_address, desig_no, city_code,
                                    applicant_name, applicant_address,
                                    year, month, tax_amount):
    """宿泊税納入書テンプレ.xlsx を読み込み、3面分のデータを書き込んで返す。"""
    wb = load_workbook(_NONYUSHO_TPL)
    # シート名の前後スペースを吸収して検索
    ws = next((wb[s] for s in wb.sheetnames if s.strip() == '第12号様式'), None)
    if ws is None:
        raise KeyError(f"「第12号様式」シートが見つかりません。シート名を確認してください。現在のシート: {wb.sheetnames}")

    nm = month % 12 + 1
    ny = year + (1 if month == 12 else 0)
    due_str = f'令和{reiwa(ny)}年{nm}月{calendar.monthrange(ny, nm)[1]}日'
    _tstr   = str(max(0, int(tax_amount)))
    _padded = _tstr.zfill(9)
    _first  = next((i for i, d in enumerate(_padded) if d != '0'), 8)
    digits9 = ['' if i < _first else int(d) for i, d in enumerate(_padded)]
    addr = facility_address or applicant_address or ''

    # 3面: Panel1=col A(1), Panel2=col O(15), Panel3=col AC(29)
    for s in (1, 15, 29):
        ws.cell(row=10, column=s    ).value = f'令和{reiwa(year)}年度'  # 年度
        ws.cell(row=10, column=s+2  ).value = f'{month}月'              # 申告年月
        ws.cell(row=10, column=s+8  ).value = desig_no or ''            # 指定番号
        for i, d in enumerate(digits9):
            ws.cell(row=11, column=s+4+i).value = d                     # 税額（桁ごと）
            ws.cell(row=14, column=s+4+i).value = d                     # 合計額（桁ごと）
        ws.cell(row=15, column=s+4  ).value = due_str                   # 納期限
        ws.cell(row=18, column=s    ).value = addr                      # 住所（所在地）
        ws.cell(row=20, column=s    ).value = applicant_name or ''      # 氏名（名称）

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def generate_receipt(receipt_type, guest_name, room_no, guests,
                     stay_date, room_rate, consumption_tax,
                     issue_date, facility_name, facility_address):
    """領収書テンプレートに値を書き込んで BytesIO を返す。
    receipt_type: 'A' | 'B' | 'C' | 'D'
    """
    tpl_map = {
        'A': _RECEIPT_TPL_A,
        'B': _RECEIPT_TPL_B,
        'C': _RECEIPT_TPL_C,
        'D': _RECEIPT_TPL_D,
    }
    wb = load_workbook(tpl_map[receipt_type])
    ws = wb.active

    _, _, tax_amount = calc_tax(guests, room_rate)

    def w(row, col, val):
        ws.cell(row=row, column=col).value = val

    def yen(amount):
        return f'{int(amount):,}円'

    def fmt_date(d):
        return f'令和{reiwa(d.year)}年{d.month}月{d.day}日'

    stay_str = f'{stay_date.month}月{stay_date.day}日'

    if receipt_type == 'A':
        # タイプA: 10列, 発行日はD/F/H列に分割
        w(3, 2, guest_name or '')
        w(3, 8, room_no or '')
        w(4, 8, guests)
        w(7, 2, stay_str)
        w(7, 4, yen(room_rate))
        w(8, 4, yen(consumption_tax))
        w(9, 4, yen(tax_amount))
        w(12, 4, yen(room_rate + consumption_tax + tax_amount))
        # 発行日 1枚目 (row 15)
        w(15, 4, f'令和{reiwa(issue_date.year)}')
        w(15, 6, issue_date.month)
        w(15, 8, issue_date.day)
        w(16, 4, facility_address or '')
        w(18, 4, facility_name or '')
        # 発行日 2枚目 (row 20)
        w(20, 4, f'令和{reiwa(issue_date.year)}')
        w(20, 6, issue_date.month)
        w(20, 8, issue_date.day)
        w(21, 4, facility_address or '')
        w(22, 4, facility_name or '')

    elif receipt_type == 'B':
        # タイプB: 4列, 発行日はD列に文字列
        w(3, 2, guest_name or '')
        w(3, 4, f'{room_no}号室' if room_no else '')
        w(4, 4, f'人数{guests}名')
        w(7, 2, stay_str)
        w(7, 4, yen(room_rate))
        w(8, 4, yen(consumption_tax))
        w(11, 4, yen(room_rate + consumption_tax))
        w(13, 2, f'上記のほか、宿泊税額{yen(tax_amount)}を領収しました。')
        w(16, 4, fmt_date(issue_date))
        w(17, 4, facility_address or '')
        w(19, 4, facility_name or '')
        w(21, 4, fmt_date(issue_date))
        w(22, 4, facility_address or '')
        w(23, 4, facility_name or '')

    elif receipt_type == 'C':
        # タイプC: 4列, 総額のみ
        total = room_rate + consumption_tax + tax_amount
        w(3, 2, guest_name or '')
        w(3, 4, f'{room_no}号室' if room_no else '')
        w(4, 4, f'人数{guests}名')
        w(7, 2, stay_str)
        w(7, 4, yen(total))
        w(11, 4, yen(total))
        w(13, 2, f'上記金額には、消費税額等{yen(consumption_tax)}及び宿泊税額{yen(tax_amount)}が含まれています。')
        w(16, 4, fmt_date(issue_date))
        w(17, 4, facility_address or '')
        w(19, 4, facility_name or '')
        w(21, 4, fmt_date(issue_date))
        w(22, 4, facility_address or '')
        w(23, 4, facility_name or '')

    elif receipt_type == 'D':
        # タイプD: 宿泊税のみ, 4列
        w(4, 2, f'{guest_name}様' if guest_name else '')  # B4:C4 merged
        w(4, 4, f'{room_no}号室' if room_no else '')
        w(5, 4, f'人数　{guests}名')
        w(8, 2, stay_str)
        w(8, 4, yen(tax_amount))
        w(12, 4, yen(tax_amount))
        w(14, 2, f'上記金額は、宿泊税{yen(tax_amount)}です。')  # B14:D14 merged
        w(16, 4, fmt_date(issue_date))
        w(17, 4, facility_address or '')
        w(19, 4, facility_name or '')
        w(21, 4, fmt_date(issue_date))
        w(22, 4, facility_address or '')
        w(24, 4, facility_name or '')  # D24（D23ではない）

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────── テーマ定数 ───────────
import base64 as _b64
_logo_path = os.path.join(_BASE_DIR, '宿泊税スマートロゴ.png')
if os.path.exists(_logo_path):
    with open(_logo_path, 'rb') as _f:
        _logo_b64 = _b64.b64encode(_f.read()).decode()
    _LOGO_HTML = f'<img src="data:image/png;base64,{_logo_b64}" width="64" height="64" style="border-radius:12px;">'
else:
    _LOGO_HTML = ''

_CSS = """<style>
[data-testid="stAppViewContainer"]{background:#EBF2FA;}
[data-testid="stHeader"]{background:transparent;}
.block-container{padding-top:.8rem!important;}
[data-testid="stSidebar"]{background:linear-gradient(160deg,#1B4F72 0%,#0E2F44 100%)!important;}
[data-testid="stSidebar"] *{color:#D6EAF8!important;}
[data-testid="stSidebar"] input{background:rgba(255,255,255,.92)!important;color:#1B4F72!important;border:1px solid rgba(255,255,255,.4)!important;border-radius:6px!important;}
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
  <div style="flex-shrink:0">{_LOGO_HTML}</div>
  <div style="flex:1">
    <div style="color:white;font-size:26px;font-weight:700;letter-spacing:-.3px;line-height:1.1;margin-bottom:5px;">宿泊税スマート</div>
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
_cfg = load_settings()

with st.sidebar:
    st.header('⚙️ 施設・期間設定')
    city_code = st.text_input('市町村コード',
                               value=_cfg.get('city_code', '473260'),
                               help='北谷町：473260')
    desig_no = st.text_input('指定番号',
                              value=_cfg.get('desig_no', ''),
                              placeholder='例：001')
    facility_name = st.text_input('宿泊施設名',
                                   value=_cfg.get('facility_name', ''),
                                   placeholder='例：○○ホテル')
    facility_address = st.text_input('施設所在地',
                                      value=_cfg.get('facility_address', ''),
                                      placeholder='例：沖縄県中頭郡北谷町桑江1-1-1')
    applicant_name = st.text_input('申告者名（法人名・代表者名）',
                                    value=_cfg.get('applicant_name', ''))
    applicant_address = st.text_input('申告者住所',
                                       value=_cfg.get('applicant_address', ''))

    if st.button('💾 施設情報を保存', use_container_width=True):
        save_settings({
            'city_code':        city_code,
            'desig_no':         desig_no,
            'facility_name':    facility_name,
            'facility_address': facility_address,
            'applicant_name':   applicant_name,
            'applicant_address':applicant_address,
            'tax_inclusive':    tax_inclusive,
            'meal_per_person':  int(meal_per_person),
        })
        st.success('保存しました。次回起動時も自動入力されます。')

    st.divider()
    st.subheader('📅 申告対象年月')
    col_y, col_m = st.columns(2)
    _today = date.today()
    with col_y:
        target_year = st.number_input('年（西暦）', min_value=2025, max_value=2035,
                                       value=_cfg.get('last_year', _today.year))
    with col_m:
        target_month = st.number_input('月', min_value=1, max_value=12,
                                        value=_cfg.get('last_month', _today.month))

    st.divider()
    st.subheader('💴 料金設定')
    tax_inclusive = st.checkbox(
        '宿泊料金は消費税込み',
        value=_cfg.get('tax_inclusive', False),
        help='CSVの料金が消費税込みの場合はON。÷1.10で税抜き額に換算してから宿泊税を計算します。',
    )
    meal_per_person = st.number_input(
        '食事代控除額（円/人）',
        min_value=0, value=int(_cfg.get('meal_per_person', 0)), step=100,
        help='1泊1人あたりの食事代。課税標準額から控除されます。不要なら0のまま。',
    )

    st.divider()
    st.subheader('🎓 課税免除（学校行事等）')
    st.caption(
        '⚠️ 課税免除対象の宿泊は、先に「計算・確認」タブの日別編集で該当日の人数・金額を差し引いてください。'
        'この欄は申告書への記載用です。ここに入力しても税額の計算には影響しません。'
    )
    exempt_guests = st.number_input('課税免除 宿泊者数（延べ数）', min_value=0, value=0)
    exempt_amount = st.number_input('課税免除 宿泊料金合計（円）', min_value=0, value=0)

# ── メインタブ ──
tab_upload, tab_calc, tab_export, tab_receipt = st.tabs([
    '📁  データ読み込み',
    '📊  計算・確認',
    '📄  書類出力',
    '🧾  領収書発行',
])

# ── タブ1: CSV読み込み ──
with tab_upload:
    st.markdown(_sec('📁', '売上CSVファイルの読み込み',
                     '3ヶ月分の予約CSV（ホテル管理システム／Airbnb）をアップロードします。同じ月でも複数ファイル選択可'), unsafe_allow_html=True)

    _periods = three_month_period(int(target_year), int(target_month))
    _period_labels = [f'令和{reiwa(y)}年{m}月' for y, m in _periods]
    st.info(f'申告期間：**{_period_labels[0]}** ～ **{_period_labels[2]}**　｜　サイドバーの「申告対象年月」で最終月を変更できます')

    _cols = st.columns(3)
    _months_items = st.session_state.get('months_items', {})

    for _idx, (_col, (y, m), _label) in enumerate(zip(_cols, _periods, _period_labels)):
        with _col:
            st.markdown(f'**第{_idx+1}月：{_label}**')
            _ufs = st.file_uploader(
                f'{_label} CSV',
                type=['csv', 'CSV'],
                key=f'upload_{_idx}',
                accept_multiple_files=True,
                help='ねっぱん・AirbnbなどCSVを複数同時選択可。形式は自動判定して合算します',
            )
            if _ufs:
                _all_items  = []
                _raw_csv    = None
                _csv_format = None
                _fmt_labels = []

                _has_airbnb     = False
                _airbnb_raw_csv = None
                for _uf in _ufs:
                    _raw = _uf.read()
                    _fmt = detect_csv_format(_raw)
                    if _fmt == 'airbnb':
                        _all_items  += parse_airbnb_csv(_raw, y, m)
                        _has_airbnb = True
                        if _airbnb_raw_csv is None:
                            _airbnb_raw_csv = _raw
                        _fmt_labels.append('Airbnb')
                    elif _fmt == 'hotel_total':
                        _all_items  += parse_hotel_total_csv(_raw, y, m)
                        if _raw_csv is None:
                            _raw_csv    = _raw
                            _csv_format = 'hotel_total'
                        _fmt_labels.append('ホテルシステム(集計)')
                    else:
                        _all_items  += parse_hotel_csv(_raw, y, m)
                        if _raw_csv is None:
                            _raw_csv    = _raw
                            _csv_format = 'hotel'
                        _fmt_labels.append('ホテルシステム')

                _months_items[(y, m)] = {'items': _all_items, 'raw_csv': _raw_csv,
                                          'csv_format': _csv_format,
                                          'has_airbnb': _has_airbnb,
                                          'airbnb_raw_csv': _airbnb_raw_csv}
                st.caption('📋 ' + ' ＋ '.join(_fmt_labels))

                if _all_items:
                    _rd = {}
                    for _i in _all_items:
                        _d = _i['day']
                        _rd[_d] = _rd.get(_d, 0) + _i['guests']
                    st.success(f'✅ {len(_all_items)}件 / {sum(_rd.values())}人泊')
                else:
                    st.warning('データが見つかりません')

    st.session_state['months_items'] = _months_items

    # 後方互換: 最終月データを items に保存（計算確認タブ用）
    _last_key = (int(target_year), int(target_month))
    if _last_key in _months_items:
        st.session_state['items']      = _months_items[_last_key]['items']
        st.session_state['raw_csv']    = _months_items[_last_key]['raw_csv']
        st.session_state['csv_format'] = _months_items[_last_key].get('csv_format')
        st.session_state['has_airbnb']     = _months_items[_last_key].get('has_airbnb', False)
        st.session_state['airbnb_raw_csv'] = _months_items[_last_key].get('airbnb_raw_csv')

    # ── 読み込み済みサマリー ──
    if _months_items:
        st.divider()
        st.markdown('##### 読み込み済みデータ')
        _summary_rows = []
        for y, m in _periods:
            _k = (y, m)
            if _k in _months_items and _months_items[_k]['items']:
                _its = _months_items[_k]['items']
                _summary_rows.append({
                    '月': f'令和{reiwa(y)}年{m}月',
                    '宿泊レコード数': len(_its),
                    '延べ宿泊者数': sum(i['guests'] for i in _its),
                    '売上合計（円）': f"{sum(i['amount'] for i in _its):,}",
                })
            else:
                _summary_rows.append({'月': f'令和{reiwa(y)}年{m}月', '宿泊レコード数': '─', '延べ宿泊者数': '─', '売上合計（円）': '─'})
        st.dataframe(pd.DataFrame(_summary_rows), use_container_width=True, hide_index=True)
        st.caption('※ 月をまたぐ滞在は宿泊日で自動分割し、対象月分のみを集計しています。')

# ── タブ2: 計算結果 ──
with tab_calc:
    st.markdown(_sec('📊', '宿泊税 計算結果',
                     '日別の課税標準額・宿泊税額を確認します'), unsafe_allow_html=True)

    if 'items' not in st.session_state or not st.session_state['items']:
        st.info('「データ読み込み」タブでCSVファイルをアップロードしてください。')
    else:
        items = st.session_state['items']

        # CSV から日別集計（元データ）
        _daily_csv = compute_daily_agg(
            items,
            meal_per_person=int(meal_per_person),
            tax_inclusive=tax_inclusive,
        )

        # ── 日別データ 手動修正テーブル ──
        st.markdown(_sec('✏️', '日別データの確認・修正',
                         '宿泊者数や宿泊料金がCSVと実際で異なる場合は直接編集してください', '#E67E22'),
                    unsafe_allow_html=True)
        st.caption('💡 セルをクリックして直接入力できます。CSVを再アップロードするとリセットされます。')

        # CSVが変わったらエディタ状態をリセット
        _fp      = str(sorted((i['day'], i['guests'], i['amount']) for i in items))
        _fp_key  = f'csv_fp_{target_year}_{target_month}'
        _edt_key = f'daily_editor_{target_year}_{target_month}'
        if st.session_state.get(_fp_key) != _fp:
            st.session_state[_fp_key] = _fp
            st.session_state.pop(_edt_key, None)

        # 初期データを構築（エディタリセット時のみ使われる）
        _init_rows = [
            {
                '日付':           f'{target_month}/{d}',
                '宿泊者数':       _daily_csv[d]['guests'],
                '宿泊料金合計（円）': _daily_csv[d]['amount'],
            }
            for d in sorted(_daily_csv.keys())
        ]
        _init_df = pd.DataFrame(_init_rows)

        _edited_df = st.data_editor(
            _init_df,
            key=_edt_key,
            column_config={
                '日付':              st.column_config.TextColumn('日付', disabled=True),
                '宿泊者数':          st.column_config.NumberColumn('宿泊者数', min_value=0, step=1),
                '宿泊料金合計（円）': st.column_config.NumberColumn('宿泊料金合計（円）', min_value=0, step=100, format='%d'),
            },
            hide_index=True,
            use_container_width=True,
        )

        # 日別集計の確定：未編集日は予約別計算（正確）、編集日のみ日次平均で再計算
        daily_data = dict(_daily_csv)  # 予約別計算結果を基本値とする
        for _, _row in _edited_df.iterrows():
            try:
                _d = int(str(_row['日付']).split('/')[-1])
            except (ValueError, TypeError):
                continue
            _g = max(0, int(_row.get('宿泊者数', 0) or 0))
            _a = max(0, int(_row.get('宿泊料金合計（円）', 0) or 0))
            _orig = _daily_csv.get(_d)
            if _orig and _orig['guests'] == _g and _orig['amount'] == _a:
                continue  # 未編集 → 予約別計算結果を維持
            # 手動編集あり → 日次合計から再計算（やむを得ず平均割り）
            if _g > 0 and _a > 0:
                _adj  = (_a / 1.10) if tax_inclusive else _a
                _net  = _adj - int(meal_per_person) * _g
                _pp   = _net / _g
                _bp   = math.floor(_pp / 1000) * 1000
                _tpp  = min(_bp * TAX_RATE, MAX_TAX_PER_PERSON)
                _base = int(_bp * _g)
                _tax  = int(_tpp * _g)
            else:
                _base = _tax = 0
            daily_data[_d] = {'guests': _g, 'base': _base, 'tax': _tax, 'amount': _a}

        # 書類出力タブ用に編集済みデータを保存
        _months_daily = st.session_state.get('months_daily', {})
        _months_daily[(int(target_year), int(target_month))] = daily_data
        st.session_state['months_daily'] = _months_daily

        total_taxable_guests = sum(v['guests'] for v in daily_data.values())
        total_taxable_base   = sum(v['base']   for v in daily_data.values())
        total_tax            = sum(v['tax']    for v in daily_data.values())

        st.session_state['calc'] = {
            'total_guests': total_taxable_guests,
            'total_base': total_taxable_base,
            'total_tax': total_tax,
        }
        st.session_state['daily_data'] = daily_data

        # 最後に使った年月を settings に保存（次回起動時のデフォルトにする）
        if os.path.exists(_SETTINGS_FILE):
            try:
                _cur = load_settings()
                _cur.update({'last_year': int(target_year), 'last_month': int(target_month),
                             'tax_inclusive': tax_inclusive, 'meal_per_person': int(meal_per_person)})
                save_settings(_cur)
            except Exception:
                pass

        st.divider()

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
            _has_hotel_raw  = bool(st.session_state.get('raw_csv'))
            _has_airbnb_raw = bool(st.session_state.get('airbnb_raw_csv'))

            if not _has_hotel_raw and not _has_airbnb_raw:
                st.info('「データ読み込み」タブでCSVを再アップロードすると検証データが表示されます。')
            else:
                # ── Airbnb 予約別明細を表示するヘルパー ──
                def _show_airbnb_detail(prefix=''):
                    ab_detail, ab_excluded = parse_airbnb_csv_for_verification(
                        st.session_state['airbnb_raw_csv'],
                        int(target_year), int(target_month),
                        meal_per_person=int(meal_per_person),
                        tax_inclusive=tax_inclusive,
                    )
                    st.subheader(f'{prefix}Airbnb 予約別 計算ステップ明細')
                    st.caption(
                        '**課税標準額/人** = 1泊料金 ÷ 人数 を千円未満切捨　'
                        '**宿泊税** = 課税標準額/人 × 2% × 人数（上限2,000円/人）'
                    )
                    if ab_detail:
                        df_ab = pd.DataFrame(ab_detail)
                        for col in ['1泊料金（円）', '課税標準額/人（千円切捨）', '宿泊税（円）']:
                            if col in df_ab.columns:
                                df_ab[col] = df_ab[col].apply(
                                    lambda x: f'{x:,}' if isinstance(x, (int, float)) else x)
                        st.dataframe(df_ab, use_container_width=True, hide_index=True)
                        total_ab_tax = sum(r['宿泊税（円）'] for r in ab_detail)
                        st.caption(f'Airbnb 宿泊税合計: ¥{total_ab_tax:,}')
                    else:
                        st.info('対象月のAirbnb予約データがありません。')

                    st.divider()
                    if ab_excluded:
                        st.subheader(f'{prefix}Airbnb 月またぎ除外データ（{len(ab_excluded)} 件）')
                        st.caption(
                            f'CSVに含まれていたが **{target_month}月以外の宿泊夜** のため計算から除外したデータです。'
                        )
                        st.dataframe(pd.DataFrame(ab_excluded), use_container_width=True, hide_index=True)
                    else:
                        st.subheader(f'{prefix}Airbnb 月またぎ除外データ')
                        st.success(f'月をまたぐAirbnb宿泊データはありません。すべて {target_month}月内の宿泊です。')
                    return ab_detail, ab_excluded

                # ── Airbnbのみの場合 ──
                if not _has_hotel_raw:
                    _show_airbnb_detail()

                # ── ホテルCSV（集計形式）の場合 ──
                elif st.session_state.get('csv_format') == 'hotel_total':
                    st.info('ℹ️ 集計形式CSV（ReservationTotalList）では日別照合は利用できません。日別明細形式のCSVをご利用ください。')
                    if _has_airbnb_raw:
                        st.divider()
                        _show_airbnb_detail()

                # ── ホテルCSV（日別明細形式）の場合 ──
                else:
                    if _has_airbnb_raw:
                        st.info('ℹ️ ホテルシステムCSVとAirbnb CSVが混在しています。それぞれ個別に検証データを表示します。')
                        st.subheader('【ホテルシステムCSV】')

                    # ── 日別集計 ＋ 予約別明細 ＋ 除外データ ＋ Agoda補正 ──
                    v_detail, v_sub, v_excluded, v_agoda = parse_hotel_csv_for_verification(
                        st.session_state['raw_csv'],
                        int(target_year), int(target_month),
                        meal_per_person=int(meal_per_person),
                        tax_inclusive=tax_inclusive,
                    )

                    st.subheader('① 日別集計（予約データから算出）')
                    st.caption('チェックイン〜チェックアウトの日付をもとに1日単位で集計した結果です。')
                    comp_rows = []
                    for d in sorted(set(list(daily_data.keys()) + list(v_sub.keys()))):
                        our_g = daily_data.get(d, {}).get('guests', 0)
                        our_b = daily_data.get(d, {}).get('base', 0)
                        our_t = daily_data.get(d, {}).get('tax', 0)
                        v_a   = v_sub.get(d, {}).get('amount', 0)
                        comp_rows.append({
                            '日付':               f'{target_month}/{d}',
                            '人数（人）':          our_g,
                            '売上合計（円）':       f'{v_a:,}',
                            '課税標準額合計（円）': f'{our_b:,}',
                            '宿泊税（円）':         f'{our_t:,}',
                        })
                    st.dataframe(pd.DataFrame(comp_rows), use_container_width=True, hide_index=True)

                    st.divider()
                    st.subheader('② 予約別 計算ステップ明細')
                    st.caption(
                        '**課税標準額/人** = 1泊料金 ÷ 人数 を千円未満切捨　'
                        '**宿泊税** = 課税標準額/人 × 2% × 人数（上限2,000円/人）'
                    )

                    if v_detail:
                        df_v = pd.DataFrame(v_detail)
                        num_cols = ['1泊料金（円）', '基本利用（円）', '1人当たり料金',
                                    '課税標準額/人（千円切捨）', '課税標準額合計', '宿泊税（円）']
                        for col in num_cols:
                            if col in df_v.columns:
                                df_v[col] = df_v[col].apply(lambda x: f'{x:,}' if isinstance(x, (int, float)) and x != '' else x)
                        st.dataframe(df_v, use_container_width=True, hide_index=True)

                    # ── ③ 月またぎ除外データ ──
                    st.divider()
                    if v_excluded:
                        st.subheader(f'③ 月またぎ除外データ（{len(v_excluded)} 件）')
                        st.caption(
                            f'CSVに含まれていたが **{target_month}月以外の宿泊夜** のため計算から除外したデータです。'
                            '月をまたぐ予約では、対象月外の泊数分が自動的に除かれています。'
                        )
                        df_ex = pd.DataFrame(v_excluded)
                        st.dataframe(df_ex, use_container_width=True, hide_index=True)
                    else:
                        st.subheader('③ 月またぎ除外データ')
                        st.success(f'月をまたぐ宿泊データはありません。すべて {target_month}月内の宿泊です。')

                    # ── ④ Agoda補正明細 ──
                    st.divider()
                    if v_agoda:
                        agoda_total_raw   = sum(r['CSV記載額（控除後）'] for r in v_agoda)
                        agoda_total_gross = sum(r['グロスアップ後（円）'] for r in v_agoda)
                        st.subheader(f'④ Agoda補正明細（{len(v_agoda)} 件）')
                        st.caption(
                            f'AgodaはCSVに **手数料控除後の金額** が記載されます。'
                            f'通常: 12%（÷0.88）、Agoda Intelligence チャネル: 9%（÷0.91）でグロスアップしています。'
                            f'　補正合計: ¥{agoda_total_raw:,} → ¥{agoda_total_gross:,}'
                            f'（差額 ¥{agoda_total_gross - agoda_total_raw:,}）'
                        )
                        df_ag = pd.DataFrame(v_agoda)
                        for col in ['CSV記載額（控除後）', 'グロスアップ後（円）', '補正額（手数料分）']:
                            if col in df_ag.columns:
                                df_ag[col] = df_ag[col].apply(lambda x: f'{x:,}' if isinstance(x, int) else x)
                        st.dataframe(df_ag, use_container_width=True, hide_index=True)
                    else:
                        st.subheader('④ Agoda補正明細')
                        st.success('Agoda予約はありません。補正なし。')

                    # ── Airbnb混在時：Airbnb検証を追加表示 ──
                    if _has_airbnb_raw:
                        st.divider()
                        st.subheader('【Airbnb CSV】')
                        _show_airbnb_detail()

                    if v_detail:
                        # 検証Excelダウンロード（4シート構成）
                        def _make_verify_excel(detail, comparison, excluded, agoda):
                            def _write_sheet(ws, rows, title_color='DDEEFF'):
                                if not rows:
                                    ws.cell(row=1, column=1, value='データなし')
                                    return
                                cols = list(rows[0].keys())
                                for ci, h in enumerate(cols, 1):
                                    c = ws.cell(row=1, column=ci, value=h)
                                    c.font = Font(bold=True, size=9, name='MS Gothic')
                                    c.fill = PatternFill(fill_type='solid', fgColor=title_color)
                                    c.border = thin_border()
                                    c.alignment = Alignment(horizontal='center', vertical='center')
                                for ri, row in enumerate(rows, 2):
                                    for ci, key in enumerate(cols, 1):
                                        c = ws.cell(row=ri, column=ci, value=row[key])
                                        c.font = Font(size=9, name='MS Gothic')
                                        c.border = thin_border()
                                        c.alignment = Alignment(horizontal='center', vertical='center')
                                for col in ws.columns:
                                    ws.column_dimensions[col[0].column_letter].width = 18

                            wb  = Workbook()
                            ws1 = wb.active
                            ws1.title = '②予約別計算明細'
                            _write_sheet(ws1, detail, 'DDEEFF')

                            ws2 = wb.create_sheet('①日別集計')
                            _write_sheet(ws2, comparison, 'DDEEFF')

                            ws3 = wb.create_sheet('③月またぎ除外')
                            _write_sheet(ws3, excluded if excluded else [{'内容': '除外データなし'}], 'FFE5CC')

                            ws4 = wb.create_sheet('④Agoda補正')
                            _write_sheet(ws4, agoda if agoda else [{'内容': 'Agoda補正なし'}], 'E8F5E9')

                            buf = io.BytesIO()
                            wb.save(buf)
                            buf.seek(0)
                            return buf

                        vbuf = _make_verify_excel(v_detail, comp_rows, v_excluded, v_agoda)
                        st.download_button(
                            label='📥 検証Excelをダウンロード（4シート：計算明細・日別集計・除外データ・Agoda補正）',
                            data=vbuf,
                            file_name=f'宿泊税検証_{target_year}{target_month:02d}.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        )
                    else:
                        st.info('対象月の予約データがありません。')

# ── タブ3: 書類出力 ──
with tab_export:
    st.markdown(_sec('📄', '書類出力', '月計表（3シート）・申告書（3ヶ月）・納入書（3ヶ月合計）を出力します'), unsafe_allow_html=True)

    _exp_periods  = three_month_period(int(target_year), int(target_month))
    _exp_mi       = st.session_state.get('months_items', {})

    # 各月の計算を実行（計算確認タブで手動修正済みのデータを優先使用）
    _months_daily = st.session_state.get('months_daily', {})
    _exp_months_calc = []
    for _ey, _em in _exp_periods:
        _ek = (_ey, _em)
        if _ek in _exp_mi and _exp_mi[_ek]['items']:
            _its = _exp_mi[_ek]['items']
            if _ek in _months_daily:
                _dd = _months_daily[_ek]   # 手動修正済みデータを使用
            else:
                _dd = compute_daily_agg(_its, meal_per_person=int(meal_per_person), tax_inclusive=tax_inclusive)
            _c   = {
                'year': _ey, 'month': _em,
                'total_guests': sum(v['guests'] for v in _dd.values()),
                'total_base':   sum(v['base']   for v in _dd.values()),
                'total_tax':    sum(v['tax']     for v in _dd.values()),
                'daily_data':   _dd,
            }
            _exp_months_calc.append(_c)
        else:
            _exp_months_calc.append(None)

    _loaded = [c for c in _exp_months_calc if c is not None]

    if not _loaded:
        st.info('「データ読み込み」タブで少なくとも1ヶ月分のCSVをアップロードしてください。')
    else:
        # 合計の表示
        _total_tax_3m = sum(c['total_tax'] for c in _loaded)
        _period_str   = f'令和{reiwa(_exp_periods[0][0])}年{_exp_periods[0][1]}月 ～ 令和{reiwa(_exp_periods[2][0])}年{_exp_periods[2][1]}月'
        st.info(f'申告期間：**{_period_str}**　｜　{len(_loaded)}ヶ月読込済　｜　3ヶ月合計宿泊税：**¥{_total_tax_3m:,}**')

        # 月別サマリー
        _sum_rows = []
        for c in _exp_months_calc:
            if c:
                _sum_rows.append({'月': f'令和{reiwa(c["year"])}年{c["month"]}月', '宿泊者数': f'{c["total_guests"]:,}人',
                                   '課税標準額': f'¥{c["total_base"]:,}', '宿泊税額': f'¥{c["total_tax"]:,}'})
            else:
                _lbl = next(f'令和{reiwa(y)}年{m}月' for (y,m), mc in zip(_exp_periods, _exp_months_calc) if mc is None)
                _sum_rows.append({'月': _lbl, '宿泊者数': '─', '課税標準額': '─', '宿泊税額': '─'})
        st.dataframe(pd.DataFrame(_sum_rows), use_container_width=True, hide_index=True)

        col_a, col_b, col_c = st.columns(3)

        with col_a:
            st.markdown(_dl_card('①', '宿泊税月計表',
                                 '申告書に添付する月別日次計表',
                                 f'{len(_loaded)}ヶ月分を1ファイル・複数シートで出力', '📋'),
                        unsafe_allow_html=True)
            if os.path.exists(_MONTHLY_TPL):
                _monthly_buf = generate_monthly_from_template(
                    facility_name=facility_name or '（未入力）',
                    desig_no=desig_no or '―',
                    months_calc_list=_loaded,
                )
            else:
                _wb_monthly = Workbook()
                if 'Sheet' in _wb_monthly.sheetnames:
                    del _wb_monthly['Sheet']
                for _mc in _loaded:
                    generate_monthly_table(
                        facility_name=facility_name or '（未入力）',
                        desig_no=desig_no or '―',
                        city_code=city_code or '―',
                        year=_mc['year'], month=_mc['month'],
                        taxable_guests=_mc['total_guests'],
                        taxable_amount=_mc['total_base'],
                        tax_amount=_mc['total_tax'],
                        daily_data=_mc['daily_data'],
                        exempt_guests=exempt_guests, exempt_amount=exempt_amount,
                        _workbook=_wb_monthly,
                    )
                _monthly_buf = io.BytesIO()
                _wb_monthly.save(_monthly_buf)
                _monthly_buf.seek(0)
            _fname_m = f'{_exp_periods[0][0]}{_exp_periods[0][1]:02d}-{_exp_periods[2][0]}{_exp_periods[2][1]:02d}'
            st.download_button(
                label='📥 月計表をダウンロード（Excel・3シート）',
                data=_monthly_buf,
                file_name=f'宿泊税月計表_{_fname_m}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
            )

        with col_b:
            st.markdown(_dl_card('②', '宿泊税納入申告書',
                                 '北谷町税務課に提出する申告書',
                                 f'3ヶ月分を1枚にまとめて出力', '📑'),
                        unsafe_allow_html=True)
            _decl_months = [
                {
                    'year': _mc['year'], 'month': _mc['month'],
                    'taxable_guests': _mc['total_guests'],
                    'taxable_amount': _mc['total_base'],
                    'tax_amount':     _mc['total_tax'],
                    'exempt_guests':  exempt_guests,
                    'exempt_amount':  exempt_amount,
                }
                for _mc in _loaded
            ]
            if os.path.exists(_DECLARATION_TPL):
                _decl_buf = generate_declaration_from_template(
                    facility_name=facility_name or '（未入力）',
                    facility_address=facility_address or '',
                    desig_no=desig_no or '―',
                    applicant_name=applicant_name or '',
                    applicant_address=applicant_address or '',
                    months_data=_decl_months,
                )
            else:
                _mc0 = _loaded[0]
                _decl_buf = generate_declaration(
                    facility_name=facility_name or '（未入力）',
                    facility_address=facility_address or '',
                    desig_no=desig_no or '―',
                    applicant_name=applicant_name or '',
                    applicant_address=applicant_address or '',
                    year=_mc0['year'], month=_mc0['month'],
                    taxable_guests=_mc0['total_guests'],
                    taxable_amount=_mc0['total_base'],
                    tax_amount=_mc0['total_tax'],
                    exempt_guests=exempt_guests, exempt_amount=exempt_amount,
                )
            st.download_button(
                label='📥 申告書をダウンロード（Excel）',
                data=_decl_buf,
                file_name=f'宿泊税納入申告書_{_fname_m}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
            )

        with col_c:
            st.markdown(_dl_card('③', '宿泊税納入書',
                                 '金融機関での納付に使う3枚綴り',
                                 f'3ヶ月合計税額：¥{_total_tax_3m:,}', '🏦'),
                        unsafe_allow_html=True)
            _nony_kw = dict(
                facility_name=facility_name or '（未入力）',
                facility_address=facility_address or '',
                desig_no=desig_no or '―',
                city_code=city_code or '473260',
                applicant_name=applicant_name or '',
                applicant_address=applicant_address or '',
                year=target_year, month=target_month,
                tax_amount=_total_tax_3m,
            )
            _nonyusho_buf = (generate_nonyusho_from_template(**_nony_kw)
                             if os.path.exists(_NONYUSHO_TPL)
                             else generate_nonyusho(**_nony_kw))
            st.download_button(
                label='📥 納入書をダウンロード（Excel）',
                data=_nonyusho_buf,
                file_name=f'宿泊税納入書_{_fname_m}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                type='primary',
            )

        st.divider()
        st.markdown("""
<div style="background:white;border-radius:12px;padding:18px 24px;box-shadow:0 2px 8px rgba(27,79,114,.07);">
<div style="font-weight:700;color:#1B4F72;margin-bottom:10px;font-size:14px;">📮 提出方法</div>
<ul style="margin:0;padding-left:20px;color:#5D6D7E;font-size:13px;line-height:1.9;">
<li>月計表（3シート）と申告書をセットで北谷町税務課へ提出</li>
<li>提出先：窓口持参 ／ 郵送 ／ eLTAX 利用可</li>
<li>電子申告は eLTAX（PCdeskNext）から申告可能</li>
</ul>
</div>
""", unsafe_allow_html=True)

# ── タブ4: 領収書発行 ──
with tab_receipt:
    st.markdown(_sec('🧾', '領収書発行', '宿泊税を自動計算してExcel領収書を出力します'), unsafe_allow_html=True)

    RECEIPT_LABELS = {
        'A': 'タイプA ─ 宿泊税を行ごとに記載（税込み明細）',
        'B': 'タイプB ─ 宿泊税を別途テキストで記載',
        'C': 'タイプC ─ 客室料金に全込み（内訳注記あり）',
        'D': 'タイプD ─ 宿泊税のみ（税専用領収書）',
    }

    col_r1, col_r2 = st.columns(2)

    with col_r1:
        st.markdown('##### 宿泊情報')
        r_type = st.radio('領収書タイプ', options=['A', 'B', 'C', 'D'],
                          format_func=lambda x: RECEIPT_LABELS[x])
        r_guest_name = st.text_input('氏名（宛名）', placeholder='山田 太郎')
        r_room_no    = st.text_input('部屋番号', placeholder='101')
        r_guests     = st.number_input('人数（宿泊税計算に使用）', min_value=1, max_value=99, value=1, step=1)
        r_stay_date  = st.date_input('宿泊日', value=date.today(), key='r_stay_date')

    with col_r2:
        st.markdown('##### 金額')
        r_room_rate = st.number_input('客室料金（消費税込・宿泊税除く）', min_value=0, value=11000, step=1000)
        r_cons_tax  = st.number_input('消費税等', min_value=0,
                                       value=round(r_room_rate * 10 / 110),
                                       step=1,
                                       help='税込料金の10/110を自動計算します（逆算）。変更も可能です。')
        r_room_excl = r_room_rate - r_cons_tax  # 消費税抜き額（宿泊税の課税基準）
        _, _, r_tax = calc_tax(r_guests, r_room_excl)
        st.metric('宿泊税（自動計算）', f'¥{r_tax:,}',
                  help='（税込料金 − 消費税）÷ 人数から北谷町税率で計算')
        r_issue_date = st.date_input('発行日', value=date.today(), key='r_issue_date')

    if facility_name:
        st.caption(f'施設名: **{facility_name}**　住所: {facility_address or "（未入力）"}')
    else:
        st.warning('左パネルの「宿泊施設名」を入力してください（領収書に反映されます）。')

    st.divider()

    try:
        r_buf = generate_receipt(
            receipt_type=r_type,
            guest_name=r_guest_name,
            room_no=r_room_no,
            guests=r_guests,
            stay_date=r_stay_date,
            room_rate=r_room_excl,
            consumption_tax=r_cons_tax,
            issue_date=r_issue_date,
            facility_name=facility_name or '（施設名未入力）',
            facility_address=facility_address or '',
        )
        fname_date = r_stay_date.strftime('%Y%m%d')
        fname_guest = r_guest_name.replace(' ', '_') if r_guest_name else '受取人'
        st.download_button(
            label='📥 領収書をダウンロード（Excel）',
            data=r_buf,
            file_name=f'領収書_{fname_guest}_{fname_date}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            type='primary',
            use_container_width=True,
        )
    except Exception as e:
        st.error(f'領収書生成エラー: {e}')
