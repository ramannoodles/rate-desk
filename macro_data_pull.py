"""
Macro Signal Monitor — Free Data Pull to Excel
================================================
Pulls all freely available indicators from the JPM/42M analysis.
Outputs a formatted Excel workbook with one tab per category.

REQUIREMENTS
------------
pip install fredapi yfinance pandas openpyxl requests lxml

SETUP
-----
1. Get a free FRED API key at: https://fred.stlouisfed.org/docs/api/api_key.html
2. Replace FRED_API_KEY below with your key.

RUN
---
python macro_data_pull.py
Output: macro_signals_YYYYMMDD.xlsx
"""

import os
import io
import requests
import warnings
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
from fredapi import Fred
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ── CONFIG ────────────────────────────────────────────────────────────────────
FRED_API_KEY = "36e3c1be868814301a29a85cdec4bb1f"   # get free at fred.stlouisfed.org
START_DATE   = "2015-01-01"
END_DATE     = datetime.today().strftime("%Y-%m-%d")
OUTPUT_FILE  = f"macro_signals_{datetime.today().strftime('%Y%m%d')}.xlsx"

fred = Fred(api_key=FRED_API_KEY)

# ── STYLE HELPERS ─────────────────────────────────────────────────────────────
HEADER_FILL   = PatternFill("solid", fgColor="1F3864")   # dark navy
SUB_HDR_FILL  = PatternFill("solid", fgColor="2F5496")   # medium blue
ALT_ROW_FILL  = PatternFill("solid", fgColor="EBF0FA")   # light blue
WHITE_FILL    = PatternFill("solid", fgColor="FFFFFF")
GREEN_FILL    = PatternFill("solid", fgColor="E2EFDA")
AMBER_FILL    = PatternFill("solid", fgColor="FFF2CC")
RED_FILL      = PatternFill("solid", fgColor="FFE7E7")

HEADER_FONT   = Font(name="Arial", bold=True, color="FFFFFF", size=9)
SUB_HDR_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=9)
BODY_FONT     = Font(name="Arial", size=9)
BOLD_FONT     = Font(name="Arial", bold=True, size=9)
LABEL_FONT    = Font(name="Arial", size=8, color="595959")

THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CENTER = Alignment(horizontal="center", vertical="center")
LEFT   = Alignment(horizontal="left",   vertical="center")
RIGHT  = Alignment(horizontal="right",  vertical="center")

def style_header(ws, row, cols, text, fill=HEADER_FILL, font=HEADER_FONT, height=18):
    ws.merge_cells(start_row=row, start_column=1,
                   end_row=row,   end_column=cols)
    cell = ws.cell(row=row, column=1)
    cell.value       = text
    cell.font        = font
    cell.fill        = fill
    cell.alignment   = CENTER
    ws.row_dimensions[row].height = height

def apply_table_header(ws, row, cols_list):
    for c, label in enumerate(cols_list, 1):
        cell = ws.cell(row=row, column=c)
        cell.value     = label
        cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=8)
        cell.fill      = SUB_HDR_FILL
        cell.alignment = CENTER
        cell.border    = BORDER
    ws.row_dimensions[row].height = 14

def write_data_row(ws, row_num, values, alt=False):
    fill = ALT_ROW_FILL if alt else WHITE_FILL
    for c, v in enumerate(values, 1):
        cell = ws.cell(row=row_num, column=c)
        cell.value     = v
        cell.font      = BODY_FONT
        cell.fill      = fill
        cell.border    = BORDER
        if isinstance(v, float):
            cell.alignment = RIGHT
        else:
            cell.alignment = LEFT

def auto_width(ws, min_w=8, max_w=40):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = min(max_w, max(min_w, max_len + 2))

def safe_fred(series_id, start=START_DATE, end=END_DATE):
    """Pull FRED series; return None on failure."""
    try:
        s = fred.get_series(series_id, observation_start=start,
                            observation_end=end)
        s.name = series_id
        return s
    except Exception as e:
        print(f"  [WARN] FRED {series_id}: {e}")
        return None

def safe_yf(ticker, start=START_DATE, end=END_DATE, col="Close"):
    """Pull yfinance ticker; return None on failure."""
    try:
        df = yf.download(ticker, start=start, end=end, progress=False, auto_adjust=True)
        if df.empty:
            return None
        s = df[col].squeeze()
        s.name = ticker
        return s
    except Exception as e:
        print(f"  [WARN] yfinance {ticker}: {e}")
        return None

def download_csv(url, **kwargs):
    """Download a CSV from a URL; return DataFrame or None."""
    try:
        resp = requests.get(url, timeout=15)
        resp.raise_for_status()
        return pd.read_csv(io.StringIO(resp.text), **kwargs)
    except Exception as e:
        print(f"  [WARN] CSV download {url}: {e}")
        return None


# ══════════════════════════════════════════════════════════════════════════════
# DATA PULLS — ORGANISED BY CATEGORY
# ══════════════════════════════════════════════════════════════════════════════

def pull_macro():
    print("Pulling: Macro / Growth...")
    series = {
        "Real GDP (SAAR $bn)"           : safe_fred("GDPC1"),
        "Nominal GDP (SAAR $bn)"        : safe_fred("GDP"),
        "Real GDP ex-Govt (proxy)"      : None,               # calculated below
        "Nonfarm Payrolls (000s)"        : safe_fred("PAYEMS"),
        "Private Payrolls (000s)"        : safe_fred("USPRIV"),
        "Mfg Payrolls (000s)"            : safe_fred("MANEMP"),
        "Govt Payrolls (000s)"           : safe_fred("CES9091000001"),
        "Unemployment U-3 (%)"          : safe_fred("UNRATE"),
        "Unemployment U-6 (%)"          : safe_fred("U6RATE"),
        "Initial Jobless Claims (SA)"   : safe_fred("ICSA"),
        "Continuing Claims (SA)"        : safe_fred("CCSA"),
        "JOLTS Job Openings (000s)"     : safe_fred("JTSJOL"),
        "JOLTS Hires Rate (%)"          : safe_fred("JTSHIR"),
        "JOLTS Quits Rate (%)"          : safe_fred("JTSQUR"),
        "JOLTS Layoffs Rate (%)"        : safe_fred("JTSLDR"),
        "Avg Hourly Earnings YoY (%)"   : safe_fred("CES0500000003"),
        "Avg Weekly Hours"              : safe_fred("AWHAETP"),
        "Personal Savings Rate (%)"     : safe_fred("PSAVERT"),
        "Real Disposable Income ($bn)"  : safe_fred("DSPIC96"),
        "Real PCE ($bn)"                : safe_fred("PCECC96"),
        "Long-term Unemployed (000s)"   : safe_fred("UEMPMEAN"),
    }
    df = pd.DataFrame({k: v for k, v in series.items() if v is not None})
    return df

def pull_inflation():
    print("Pulling: Inflation...")
    series = {
        "Core PCE YoY — Fed Target"     : safe_fred("PCEPILFE"),
        "PCE Headline"                  : safe_fred("PCEPI"),
        "CPI Headline (NSA)"            : safe_fred("CPIAUCSL"),
        "Core CPI (ex Food & Energy)"   : safe_fred("CPILFESL"),
        "CPI Shelter"                   : safe_fred("CUSR0000SAH1"),
        "CPI Services ex-Energy"        : safe_fred("CUSR0000SASLE"),
        "CPI Food"                      : safe_fred("CPIUFDSL"),
        "CPI Energy"                    : safe_fred("CPIENGSL"),
        "PPI Final Demand"              : safe_fred("PPIACO"),
        "10yr Breakeven Inflation (%)"  : safe_fred("T10YIE"),
        "5yr Breakeven Inflation (%)"   : safe_fred("T5YIE"),
        "5yr5yr Fwd Inflation Swap"     : safe_fred("T5YIFR"),
        "Zillow Rent ZORI (proxy)"      : None,  # pulled separately below
    }
    df = pd.DataFrame({k: v for k, v in series.items() if v is not None})

    # Zillow ZORI — free direct download
    print("  Pulling Zillow ZORI...")
    zillow_url = (
        "https://files.zillowstatic.com/research/public_csvs/zori/"
        "Metro_zori_uc_sfrcondomfr_sm_sa_month.csv"
    )
    z = download_csv(zillow_url)
    if z is not None:
        us_row = z[z["RegionName"] == "United States"]
        if not us_row.empty:
            date_cols = [c for c in z.columns if c.startswith("20")]
            zori = us_row[date_cols].T.squeeze()
            zori.index = pd.to_datetime(zori.index)
            zori.name = "Zillow ZORI ($)"
            df = df.join(zori, how="outer")

    return df

def pull_monetary():
    print("Pulling: Monetary Policy / Rates...")
    series = {
        "Fed Funds Rate Effective (%)"  : safe_fred("DFF"),
        "10yr Treasury Yield (%)"       : safe_fred("DGS10"),
        "2yr Treasury Yield (%)"        : safe_fred("DGS2"),
        "30yr Treasury Yield (%)"       : safe_fred("DGS30"),
        "3mo Treasury Yield (%)"        : safe_fred("DTB3"),
        "Real 10yr Yield TIPS (%)"      : safe_fred("DFII10"),
        "Real 5yr Yield TIPS (%)"       : safe_fred("DFII5"),
        "10-2yr Yield Spread (bp)"      : None,  # calculated below
        "SOFR (%)"                      : safe_fred("SOFR"),
        "Fed Balance Sheet ($bn)"       : safe_fred("WALCL"),
        "Fed Treasury Holdings ($bn)"   : safe_fred("TREAST"),
        "Fed MBS Holdings ($bn)"        : safe_fred("MBST"),
        "Bank Reserves ($bn)"           : safe_fred("TOTRESNS"),
        "TGA Balance ($bn)"             : safe_fred("WTREGEN"),
        "Commercial Bank Loans ($bn)"   : safe_fred("LOANS",
                                            start=START_DATE),
        "M2 Money Supply ($bn)"         : safe_fred("M2SL"),
        "USD/JPY Spot"                  : safe_yf("JPY=X"),
        "USD/EUR Spot"                  : safe_yf("EURUSD=X"),
    }
    df = pd.DataFrame({k: v for k, v in series.items() if v is not None})
    # Yield curve spread
    t10 = safe_fred("DGS10")
    t2  = safe_fred("DGS2")
    if t10 is not None and t2 is not None:
        spread = (t10 - t2) * 100
        spread.name = "10-2yr Yield Spread (bp)"
        df["10-2yr Yield Spread (bp)"] = spread
    return df

def pull_fiscal():
    print("Pulling: Fiscal / Debt...")
    series = {
        "Federal Deficit Monthly ($mn)" : safe_fred("MTSDS133FMS"),
        "Federal Debt/GDP (%)"          : safe_fred("GFDEGDQ188S"),
        "Federal Debt Outstanding ($bn)": safe_fred("GFDEBTN"),
        "Net Interest Expense ($bn ann)": safe_fred("A091RC1Q027SBEA"),
        "Govt Receipts % GDP"           : safe_fred("FYONGDA188S"),
        "Govt Outlays % GDP"            : safe_fred("FYOSDPCE"),
        "Customs Duties ($mn)"          : safe_fred("IITTRANTQ027S"),
    }
    df = pd.DataFrame({k: v for k, v in series.items() if v is not None})

    # Foreign Treasury Holdings — TIC data
    print("  Pulling Treasury TIC foreign holdings...")
    tic_url = "https://ticdata.treasury.gov/resource-center/data-chart-center/tic/Documents/slt_table5.csv"
    tic = download_csv(tic_url, skiprows=4, header=0)
    if tic is not None:
        tic.to_csv("/tmp/tic_raw.csv")  # save for reference

    return df

def pull_structural():
    print("Pulling: Structural / Productivity / Profits...")
    series = {
        "Nonfarm Productivity (Index)"  : safe_fred("OPHNFB"),
        "Unit Labor Costs (Index)"      : safe_fred("ULCNFB"),
        "Employment Cost Index"         : safe_fred("ECIALLCIV"),
        "Corp Profits After Tax ($bn)"  : safe_fred("CP"),
        "Corp Profits Before Tax ($bn)" : safe_fred("CPROFIT"),
        "Retained Earnings ($bn)"       : safe_fred("A455RC1Q027SBEA"),
        "Labor Share of Income (%)"     : safe_fred("PRS85006173"),
        "Real Equipment Invest ($bn)"   : safe_fred("Y033RC1Q027SBEA"),
        "Private Nonres Fixed Invest"   : safe_fred("PNFI"),
    }
    df = pd.DataFrame({k: v for k, v in series.items() if v is not None})

    # Atlanta Fed Wage Tracker — free direct download
    print("  Pulling Atlanta Fed Wage Tracker...")
    atl_url = (
        "https://www.atlantafed.org/-/media/documents/datafiles/"
        "chcs/wage-growth-tracker/wage-growth-data.xlsx"
    )
    try:
        resp = requests.get(atl_url, timeout=20)
        atl = pd.read_excel(io.BytesIO(resp.content), sheet_name=0,
                            engine="openpyxl")
        atl.to_csv("/tmp/atl_wage_tracker.csv", index=False)
        print("  Atlanta Fed Wage Tracker: saved to /tmp/atl_wage_tracker.csv")
    except Exception as e:
        print(f"  [WARN] Atlanta Fed Wage Tracker: {e}")

    return df

def pull_credit_market():
    print("Pulling: Credit Markets...")
    series = {
        "IG Credit OAS (bp)"            : safe_fred("BAMLC0A0CM"),
        "HY Credit OAS (bp)"            : safe_fred("BAMLH0A0HYM2"),
        "BB OAS (bp)"                   : safe_fred("BAMLH0A1HYBBEY"),
        "CCC OAS (bp)"                  : safe_fred("BAMLH0A3HYCEY"),
        "IG Corp Yield (%)"             : safe_fred("BAMLC0A0CMEY"),
        "HY Corp Yield (%)"             : safe_fred("BAMLH0A0HYM2EY"),
        "Agg Bond Yield (%)"            : safe_fred("BAMLCC0A0CMTRIV"),
        "30yr Mortgage Rate (%)"        : safe_fred("MORTGAGE30US"),
        "HH Debt Service Ratio (%)"     : safe_fred("TDSP"),
        "CC Delinquency Rate (%)"       : safe_fred("DRCCLACBS"),
        "Auto Loan Delinquency (%)"     : safe_fred("DRAUTOAS"),
        "Mortgage Delinquency (%)"      : safe_fred("DRSFRMACBS"),
    }
    df = pd.DataFrame({k: v for k, v in series.items() if v is not None})
    return df

def pull_global_markets():
    print("Pulling: Global Markets...")
    tickers = {
        "S&P 500"                       : "^GSPC",
        "S&P 500 EW (RSP)"              : "RSP",
        "Nasdaq 100"                    : "^NDX",
        "Russell 2000"                  : "^RUT",
        "MSCI EAFE (EFA)"               : "EFA",
        "MSCI Emerging Markets (EEM)"   : "EEM",
        "MSCI Europe (VGK)"             : "VGK",
        "MSCI Japan (EWJ)"              : "EWJ",
        "MSCI China (MCHI)"             : "MCHI",
        "WTI Crude Oil"                 : "CL=F",
        "Brent Crude Oil"               : "BZ=F",
        "Gold Futures"                  : "GC=F",
        "Silver Futures"                : "SI=F",
        "US Dollar Index (DXY)"         : "DX=F",
        "VIX Volatility"                : "^VIX",
        "iShares TLT (20yr+ Tsy)"       : "TLT",
        "iShares AGG (US Agg Bond)"     : "AGG",
        "iShares HYG (High Yield)"      : "HYG",
        "iShares LQD (IG Corp)"         : "LQD",
        "Utilities (XLU)"               : "XLU",
        "Energy (XLE)"                  : "XLE",
        "Technology (XLK)"              : "XLK",
        "Financials (XLF)"              : "XLF",
        "Industrials (XLI)"             : "XLI",
        "Consumer Disc (XLY)"           : "XLY",
        "Consumer Staples (XLP)"        : "XLP",
        "Healthcare (XLV)"              : "XLV",
        "Materials (XLB)"               : "XLB",
        "Semiconductors (SOXX)"         : "SOXX",
    }
    data = {}
    for name, ticker in tickers.items():
        s = safe_yf(ticker)
        if s is not None:
            data[name] = s
    df = pd.DataFrame(data)
    return df

def pull_consumer():
    print("Pulling: Consumer / Household...")
    series = {
        "Personal Savings Rate (%)"     : safe_fred("PSAVERT"),
        "Real PCE ($bn)"                : safe_fred("PCECC96"),
        "Real Goods PCE"                : safe_fred("DGDSRX1Q020SBEA"),
        "Real Services PCE"             : safe_fred("DSERRC1Q027SBEA"),
        "Retail Sales ($mn)"            : safe_fred("RSAFS"),
        "Retail ex-Auto ($mn)"          : safe_fred("RSXFS"),
        "Household Net Worth ($bn)"     : safe_fred("TNWBSHNO"),
        "HH Debt Service Ratio (%)"     : safe_fred("TDSP"),
        "Weekly Retail Gas ($/gal)"     : safe_fred("GASREGCOVW"),
        "Homeownership Rate (%)"        : safe_fred("RHORUSQ156N"),
        "Under-35 Homeownership (%)"    : safe_fred("RHORA34Q156N"),
        "CC Delinquency Rate (%)"       : safe_fred("DRCCLACBS"),
        "Auto Loan Delinquency (%)"     : safe_fred("DRAUTOAS"),
        "U of M Consumer Sentiment"     : safe_fred("UMCSENT"),
    }
    df = pd.DataFrame({k: v for k, v in series.items() if v is not None})
    return df

def pull_valuation_proxy():
    """
    Valuation tab. Forward P/E consensus is paid (Bloomberg/FactSet).
    Free proxies: CAPE from Shiller, TTM P/E via SPX price / FRED earnings.
    """
    print("Pulling: Valuation proxies...")
    series = {
        "10yr Breakeven (inflation adj)": safe_fred("T10YIE"),
        "Real 10yr Yield (%)"           : safe_fred("DFII10"),
        "EY vs Baa Spread (%)"          : None,   # calculated below
    }
    # Shiller CAPE — direct Excel download from Yale
    print("  Pulling Shiller CAPE from Yale...")
    cape_url = "https://img.stlouisfed.org/FRED/data/shiller.xls"
    cape_df = None
    # Shiller publishes at: http://www.econ.yale.edu/~shiller/data/ie_data.xls
    shiller_url = "http://www.econ.yale.edu/~shiller/data/ie_data.xls"
    try:
        resp = requests.get(shiller_url, timeout=20)
        cape_df = pd.read_excel(io.BytesIO(resp.content),
                                sheet_name="Data", skiprows=7,
                                engine="xlrd")
        cape_df.columns = cape_df.columns.str.strip()
        cape_df.to_csv("/tmp/shiller_cape.csv", index=False)
        print("  Shiller CAPE: saved to /tmp/shiller_cape.csv")
    except Exception as e:
        print(f"  [WARN] Shiller CAPE: {e}")

    # S&P 500 price for reference
    spx = safe_yf("^GSPC")
    if spx is not None:
        series["S&P 500 Price"] = spx

    # Baa yield for EY spread calc
    baa = safe_fred("DBAA")
    if baa is not None:
        series["Baa Corp Yield (%)"] = baa

    df = pd.DataFrame({k: v for k, v in series.items() if v is not None})

    note_rows = [
        ["PAID DATA — NOT IN THIS FILE", ""],
        ["Forward P/E Consensus", "Bloomberg Terminal ~$25K/yr | FactSet ~$12K/yr"],
        ["NTM EPS Estimates (S&P 500)", "Bloomberg / FactSet / Refinitiv"],
        ["S&P Global PMI (full detail)", "IHS Markit ~$5K+/yr — headline free at spglobal.com"],
        ["MOVE Index (bond vol)", "ICE subscription — limited free history"],
        ["Manager Dispersion Data", "Morningstar Direct ~$5K/yr"],
        ["Private Credit AUM", "Preqin / PitchBook ~$20-30K/yr"],
        ["Constituent-level valuation", "Bloomberg / FactSet required for dispersion analysis"],
    ]

    return df, note_rows


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL WRITER
# ══════════════════════════════════════════════════════════════════════════════

def write_sheet(wb, sheet_name, df, title, source_note=""):
    ws = wb.create_sheet(title=sheet_name)
    ws.freeze_panes = "B3"

    COLS = len(df.columns) + 1   # +1 for date column

    # Row 1 — title banner
    style_header(ws, 1, COLS, title, height=20)

    # Row 2 — column headers
    headers = ["Date"] + list(df.columns)
    apply_table_header(ws, 2, headers)

    # Data rows
    df_sorted = df.sort_index(ascending=False)
    for i, (idx, row) in enumerate(df_sorted.iterrows()):
        r = i + 3
        vals = [idx.strftime("%Y-%m-%d") if hasattr(idx, 'strftime') else str(idx)]
        vals += [round(v, 4) if isinstance(v, float) else v for v in row]
        write_data_row(ws, r, vals, alt=(i % 2 == 1))

    # Source note below data
    if source_note:
        last_r = len(df_sorted) + 3
        ws.cell(row=last_r + 1, column=1).value = source_note
        ws.cell(row=last_r + 1, column=1).font = LABEL_FONT

    auto_width(ws)
    return ws


def write_paid_note_sheet(wb, note_rows):
    ws = wb.create_sheet(title="Paid Sources Reference")
    style_header(ws, 1, 2, "Paid Data Sources — Not Included in This File", height=20)
    apply_table_header(ws, 2, ["Indicator", "Source / Cost"])
    for i, (k, v) in enumerate(note_rows):
        r = i + 3
        ws.cell(row=r, column=1).value = k
        ws.cell(row=r, column=1).font  = BOLD_FONT if "PAID" in k else BODY_FONT
        ws.cell(row=r, column=1).fill  = RED_FILL if "PAID" in k else WHITE_FILL
        ws.cell(row=r, column=1).border = BORDER
        ws.cell(row=r, column=2).value = v
        ws.cell(row=r, column=2).font  = BODY_FONT
        ws.cell(row=r, column=2).fill  = WHITE_FILL
        ws.cell(row=r, column=2).border = BORDER
    auto_width(ws, max_w=60)


def write_summary_sheet(wb, pull_time, series_counts):
    ws = wb.create_sheet(title="Summary", index=0)
    style_header(ws, 1, 3, "Macro Signal Monitor — Data Summary", height=22)

    ws.cell(row=2, column=1).value = f"Generated: {pull_time}"
    ws.cell(row=2, column=1).font  = LABEL_FONT
    ws.cell(row=2, column=1).alignment = LEFT

    apply_table_header(ws, 4, ["Category", "# Series", "Source"])
    rows = [
        ("Macro / Growth",          series_counts.get("macro",0),       "FRED"),
        ("Inflation",               series_counts.get("inflation",0),   "FRED + Zillow"),
        ("Monetary Policy / Rates", series_counts.get("monetary",0),    "FRED + yfinance"),
        ("Fiscal / Debt",           series_counts.get("fiscal",0),      "FRED + TIC"),
        ("Structural / Profits",    series_counts.get("structural",0),  "FRED + Atlanta Fed"),
        ("Credit Markets",          series_counts.get("credit",0),      "FRED (ICE BofA series)"),
        ("Global Markets / ETFs",   series_counts.get("global",0),      "yfinance"),
        ("Consumer / Household",    series_counts.get("consumer",0),    "FRED"),
        ("Valuation Proxies",       series_counts.get("valuation",0),   "FRED + yfinance + Shiller"),
    ]
    for i, (cat, n, src) in enumerate(rows):
        r = i + 5
        fill = ALT_ROW_FILL if i % 2 else WHITE_FILL
        for c, v in enumerate([cat, n, src], 1):
            cell = ws.cell(row=r, column=c)
            cell.value = v
            cell.font  = BODY_FONT
            cell.fill  = fill
            cell.border = BORDER
            cell.alignment = LEFT

    # Paid note
    note_r = len(rows) + 7
    ws.cell(row=note_r, column=1).value = (
        "NOTE: Forward P/E, NTM EPS, PMI detail, MOVE Index, and manager dispersion "
        "data require paid subscriptions (Bloomberg ~$25K/yr, FactSet ~$12K/yr). "
        "See 'Paid Sources Reference' tab."
    )
    ws.cell(row=note_r, column=1).font = Font(name="Arial", size=8, italic=True, color="7F7F7F")
    ws.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=3)
    ws.cell(row=note_r, column=1).alignment = Alignment(wrap_text=True)
    ws.row_dimensions[note_r].height = 30
    auto_width(ws)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    print(f"\n{'='*60}")
    print("  MACRO SIGNAL MONITOR — PULLING FREE DATA")
    print(f"  Date range: {START_DATE} to {END_DATE}")
    print(f"{'='*60}\n")

    if FRED_API_KEY == "YOUR_FRED_API_KEY_HERE":
        print("[ERROR] Please set your FRED_API_KEY in the script.")
        print("        Get a free key at: https://fred.stlouisfed.org/docs/api/api_key.html")
        return

    # Pull all datasets
    macro_df      = pull_macro()
    inflation_df  = pull_inflation()
    monetary_df   = pull_monetary()
    fiscal_df     = pull_fiscal()
    structural_df = pull_structural()
    credit_df     = pull_credit_market()
    global_df     = pull_global_markets()
    consumer_df   = pull_consumer()
    valuation_df, paid_rows = pull_valuation_proxy()

    counts = {
        "macro":      len(macro_df.columns),
        "inflation":  len(inflation_df.columns),
        "monetary":   len(monetary_df.columns),
        "fiscal":     len(fiscal_df.columns),
        "structural": len(structural_df.columns),
        "credit":     len(credit_df.columns),
        "global":     len(global_df.columns),
        "consumer":   len(consumer_df.columns),
        "valuation":  len(valuation_df.columns),
    }
    total = sum(counts.values())
    print(f"\nTotal series pulled: {total}")

    # Build workbook
    print(f"\nWriting Excel workbook: {OUTPUT_FILE}")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    write_summary_sheet(wb, datetime.now().strftime("%Y-%m-%d %H:%M"), counts)

    tab_map = [
        ("Macro",       macro_df,       "Macro & Growth — FRED"),
        ("Inflation",   inflation_df,   "Inflation — FRED + Zillow"),
        ("Monetary",    monetary_df,    "Monetary Policy & Rates — FRED + yfinance"),
        ("Fiscal",      fiscal_df,      "Fiscal & Debt — FRED + Treasury TIC"),
        ("Structural",  structural_df,  "Structural: Productivity & Profits — FRED"),
        ("Credit",      credit_df,      "Credit Markets — FRED (ICE BofA OAS series)"),
        ("Global Mkts", global_df,      "Global Markets & ETFs — yfinance"),
        ("Consumer",    consumer_df,    "Consumer & Household — FRED"),
        ("Valuation",   valuation_df,   "Valuation Proxies — FRED + Shiller + yfinance"),
    ]

    for sheet_name, df, title in tab_map:
        if not df.empty:
            write_sheet(wb, sheet_name, df, title,
                        source_note=f"Source: FRED/yfinance | Pulled: {END_DATE}")

    write_paid_note_sheet(wb, paid_rows)

    wb.save(OUTPUT_FILE)
    print(f"\n[DONE] Saved: {OUTPUT_FILE}")
    print(f"       Sheets: {[s.title for s in wb.worksheets]}")
    print(f"\nNote: Run weekly/monthly to refresh. For intraday,")
    print(f"      wrap in a scheduler (same GitHub Actions setup you")
    print(f"      already have for the KWP monitor).")


if __name__ == "__main__":
    main()