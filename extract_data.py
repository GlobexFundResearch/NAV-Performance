"""
NAV Data Extractor
Run this script AFTER opening and saving the Excel file in Excel
(so that LSEG formula values are cached in the file).
"""
import pyxlsb
import json
import sys
import io
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# ── UPDATE THIS PATH to point to the latest Excel file ──────────────────────
EXCEL_PATH = r"C:\Users\User\Desktop\Claude Cowork Test\03 NAV\GBS_LSEG_Daily_NAV_Full-Version.xlsb"
OUTPUT_PATH = r"C:\Users\User\Desktop\Claude Cowork Test\03 NAV\nav_data.json"
# ────────────────────────────────────────────────────────────────────────────

def excel_date(serial):
    if serial is None or not isinstance(serial, (int, float)):
        return None
    try:
        d = datetime.fromordinal(datetime(1899, 12, 30).toordinal() + int(serial))
        return d.strftime('%Y-%m-%d')
    except:
        return None

def is_formula(v):
    return v in ('0x2a', '0x1d', 'NULL')

def clean_nav(v):
    if is_formula(v) or v is None:
        return None
    if isinstance(v, (int, float)):
        return round(float(v), 4)
    return None

def clean_pct(v):
    if is_formula(v) or v is None:
        return None
    if isinstance(v, (int, float)):
        return round(float(v) * 100, 2)
    return None

def clean_date(v):
    if is_formula(v) or v is None:
        return None
    if isinstance(v, (int, float)):
        return excel_date(v)
    if isinstance(v, str) and len(v) > 5:
        return v
    return None

def clean_text(v):
    if is_formula(v) or v is None:
        return None
    if isinstance(v, str):
        return v.replace('\n', ' / ').strip()
    if isinstance(v, (int, float)):
        return None
    return str(v)

SHEET_CATEGORY = {
    'EQ_GB_US_EU_JP':           'Global & DMs',
    'EQ_ASIA_EM_CH_IN_VN_BRIC+LA': 'Asia, EM, VN, BRIC & LA',
    'EQ_SEC_THE':               'Sector & Thematic',
    'EQ_ETF (1)':               'ETF',
    'EQ_ETF (2)':               'ETF',
    'KTAM':                     'KTAM',
    'FI_FR':                    'Foreign Fixed Income',
}

SKIP_SHEETS = {'LIPPER_FUND_NAV_HIS', 'Code'}

all_funds = []
updated_date = None

with pyxlsb.open_workbook(EXCEL_PATH) as wb:
    for sheet_name in wb.sheets:
        if sheet_name in SKIP_SHEETS:
            continue

        is_etf = sheet_name in ('EQ_ETF (1)', 'EQ_ETF (2)')
        category = SHEET_CATEGORY.get(sheet_name, sheet_name)

        with wb.get_sheet(sheet_name) as sheet:
            all_rows = [[c.v for c in row] for row in sheet.rows()]

        # Get updated date from row index 1
        if updated_date is None:
            for v in all_rows[1]:
                if isinstance(v, str) and 'Updated' in v:
                    updated_date = v.replace('Updated as of ', '').strip()
                    break

        # Find header row
        header_idx = next(
            (i for i, row in enumerate(all_rows) if 'Fund Classification' in row),
            None
        )
        if header_idx is None:
            continue

        current_class = None

        for row in all_rows[header_idx + 1:]:
            if len(row) < 6:
                continue

            classification_raw = row[1]
            mutual_fund_raw = row[2]
            mutual_fund = clean_text(mutual_fund_raw)

            if not mutual_fund:
                continue

            classification = clean_text(classification_raw)
            if classification:
                current_class = classification

            if is_etf:
                # For ETF: some price values are formulas ('0x1d'), backup values at cols 25-31
                # Cols: [0]=mnem, [1]=class, [2]=thai_fund, [3]=fund_name, [4]=date, [5]=price,
                #       [6]=1d, [7]=1w, [8]=1m, [9]=3m, [10]=1y
                # Backup: [25]=mnem, [26]=price, [27]=1d, [28]=1w, [29]=1m, [30]=3m, [31]=1y
                date_val  = clean_date(row[4]) if len(row) > 4 else None
                nav_val   = clean_nav(row[5])  if len(row) > 5 else None
                pct1d     = clean_pct(row[6])  if len(row) > 6 else None
                pct1w     = clean_pct(row[7])  if len(row) > 7 else None
                pct1m     = clean_pct(row[8])  if len(row) > 8 else None
                pct3m     = clean_pct(row[9])  if len(row) > 9 else None
                pct1y     = clean_pct(row[10]) if len(row) > 10 else None

                # Use backup columns when main price is formula
                if nav_val is None and len(row) > 26:
                    nav_val = clean_nav(row[26])
                if pct1d is None and len(row) > 27:
                    pct1d = clean_pct(row[27])
                if pct1w is None and len(row) > 28:
                    pct1w = clean_pct(row[28])
                if pct1m is None and len(row) > 29:
                    pct1m = clean_pct(row[29])
                if pct3m is None and len(row) > 30:
                    pct3m = clean_pct(row[30])
                if pct1y is None and len(row) > 31:
                    pct1y = clean_pct(row[31])

                fund_name = clean_text(row[3]) if len(row) > 3 else None
                nav_label = 'Price'
            else:
                date_val  = clean_date(row[4]) if len(row) > 4 else None
                nav_val   = clean_nav(row[5])  if len(row) > 5 else None
                pct1d     = clean_pct(row[6])  if len(row) > 6 else None
                pct1w     = clean_pct(row[7])  if len(row) > 7 else None
                pct1m     = clean_pct(row[8])  if len(row) > 8 else None
                pct3m     = clean_pct(row[9])  if len(row) > 9 else None
                pct1y     = clean_pct(row[10]) if len(row) > 10 else None
                fund_name = clean_text(row[3]) if len(row) > 3 else None
                nav_label = 'NAV/Unit'

            all_funds.append({
                'id':             clean_text(row[0]) or mutual_fund,
                'category':       category,
                'classification': current_class,
                'mutual_fund':    mutual_fund,
                'fund_name':      fund_name,
                'latest_date':    date_val,
                'nav':            nav_val,
                'nav_label':      nav_label,
                'pct1d':          pct1d,
                'pct1w':          pct1w,
                'pct1m':          pct1m,
                'pct3m':          pct3m,
                'pct1y':          pct1y,
            })

output = {
    'updated':      updated_date or datetime.now().strftime('%Y-%m-%d'),
    'generated_at': datetime.now().isoformat(),
    'total':        len(all_funds),
    'funds':        all_funds,
}

with open(OUTPUT_PATH, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)

print(f"Done! {len(all_funds)} funds -> nav_data.json")
print(f"Updated: {updated_date}")

cats = {}
for f in all_funds:
    cats[f['category']] = cats.get(f['category'], 0) + 1
for k, v in cats.items():
    print(f"  {k}: {v}")

# Count how many have performance data
has_perf = sum(1 for f in all_funds if f['pct1d'] is not None)
has_nav  = sum(1 for f in all_funds if f['nav'] is not None)
print(f"\nFunds with NAV/Price: {has_nav}/{len(all_funds)}")
print(f"Funds with %1D data:  {has_perf}/{len(all_funds)}")
