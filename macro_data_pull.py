"""
macro_data_pull.py
====================
Pulls all 36 FRED series that macro_visualize.py's 17 chart panels are
built from, and writes them into the exact workbook structure it
expects: a "Dashboard" summary tab plus "Leading Indicators" /
"Business Activity" / "Coincident Indicators" / "Lagging Indicators"
tabs, each with FRED ID / Display Name / Unit / Signal Logic in
columns A-D and monthly values from column E onward.

This didn't exist before — macro_dashboard.xlsx was always something
you already had; this is a from-scratch replacement for whatever
originally produced it, built by reverse-engineering the exact
structure from your original file.

SETUP
-----
1. Get a free FRED API key (same one used elsewhere in this project,
   if you already have it): https://fred.stlouisfed.org/docs/api/api_key.html
2. pip install fredapi openpyxl pandas  (fredapi is the only new one —
   you already have the other two)
3. Make sure FRED_API_KEY is in your .env file (run_and_publish.sh
   already loads it; for a standalone run, export it yourself:
   export FRED_API_KEY=your_key_here)

USAGE
-----
python macro_data_pull.py                          # writes macro_dashboard.xlsx
python macro_data_pull.py -o macro_dashboard.xlsx   # same, explicit

WHY SOME SERIES ARE COMPUTED, NOT RAW
---------------------------------------
Five series display as "X YoY" or a rate on the dashboard, but FRED's
underlying series for them is a raw index level, not a %. This script
computes the standard 12-month percent change for those five itself:
CPIAUCSL, CPILFESL, CUSR0000SASLE, PCEPILFE, PPIFIS. Real GDP YoY
(A191RO1Q156NBEA) is NOT one of these — that series is already
published by FRED as a year-over-year percent change, so it's pulled
raw. GDP is also quarterly, not monthly — backfilled across each month
of its quarter, same as your original file's own notes described.
"""

import argparse
import os
from datetime import date

import pandas as pd
from fredapi import Fred
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

HISTORY_START = "2019-01-01"

# (FRED ID, category, display name, unit, signal logic, needs_yoy_transform)
SERIES = [
    ("TB3MS", "Leading Indicators", "3M Treasury Yield", "%", "Rising = tighter", False),
    ("DGS2", "Leading Indicators", "2Y Treasury Yield", "%", "Rising = tighter", False),
    ("DGS5", "Leading Indicators", "5Y Treasury Yield", "%", "Rising = tighter", False),
    ("DGS10", "Leading Indicators", "10Y Treasury Yield", "%", "Rising = tighter", False),
    ("DGS30", "Leading Indicators", "30Y Treasury Yield", "%", "Rising = tighter", False),
    ("T10Y3M", "Leading Indicators", "10Y-3M Yield Spread", "%", "Negative = inverted", False),
    ("T10Y2Y", "Leading Indicators", "10Y-2Y Yield Spread", "%", "Negative = inverted", False),
    ("MPRIME", "Leading Indicators", "Prime Rate", "%", "Rising = tighter", False),
    ("FEDFUNDS", "Leading Indicators", "Fed Funds Rate", "%", "Rising = tighter", False),
    ("M2SL", "Leading Indicators", "M2 Money Supply", "$B", "Rising = loose liquidity", False),
    ("HOUST", "Leading Indicators", "Housing Starts", "K units", "Rising = good", False),
    ("PERMIT", "Leading Indicators", "Building Permits", "K units", "Rising = good", False),
    ("NFCI", "Leading Indicators", "Chicago Fed Fin. Conditions", "Index", "Rising = tighter (inverted)", False),

    ("AMTMNO", "Business Activity", "Mfg New Orders", "$M", "Rising = good", False),
    ("DGORDER", "Business Activity", "Durable Goods Orders", "$M", "Rising = good", False),
    ("INDPRO", "Business Activity", "Industrial Production", "Index", "Rising = good", False),
    ("IPMAN", "Business Activity", "Industrial Prod: Mfg", "Index", "Rising = good", False),
    ("RSAFS", "Business Activity", "Retail Sales", "$M", "Rising = good", False),

    ("PAYEMS", "Coincident Indicators", "Nonfarm Payrolls", "K jobs", "Rising = good", False),
    ("JTSJOL", "Coincident Indicators", "JOLTS Job Openings", "K", "Rising = good", False),
    ("UNRATE", "Coincident Indicators", "Unemployment Rate", "%", "Rising = bad (inverted)", False),
    ("SAHMREALTIME", "Coincident Indicators", "Sahm Rule Recession Ind.", "Index", "Triggers at 0.50 (inverted)", False),
    ("AWHAETP", "Coincident Indicators", "Avg Weekly Hours", "hrs", "Rising = good", False),
    ("CES0500000003", "Coincident Indicators", "Avg Hourly Earnings", "$", "Rising = good", False),
    ("PI", "Coincident Indicators", "Personal Income", "$B", "Rising = good", False),
    ("UMCSENT", "Coincident Indicators", "UMich Consumer Sentiment", "Index", "Rising = good", False),
    ("PCE", "Coincident Indicators", "Personal Spending", "$B", "Rising = good", False),

    ("A191RO1Q156NBEA", "Lagging Indicators", "Real GDP YoY", "%", "Rising = good", False),  # already YoY on FRED
    ("CPIAUCSL", "Lagging Indicators", "CPI YoY", "%", "Rising = bad (inverted)", True),
    ("CPILFESL", "Lagging Indicators", "Core CPI YoY", "%", "Rising = bad (inverted)", True),
    ("CUSR0000SASLE", "Lagging Indicators", "Supercore CPI Proxy", "%", "Rising = bad (inverted)", True),
    ("PCEPILFE", "Lagging Indicators", "Core PCE Deflator", "%", "Rising = bad (inverted)", True),
    ("PPIFIS", "Lagging Indicators", "PPI Final Demand", "%", "Rising = bad (inverted)", True),
    ("T5YIE", "Lagging Indicators", "5Y Breakeven Inflation", "%", "Market inflation expectation", False),
    ("T10YIE", "Lagging Indicators", "10Y Breakeven Inflation", "%", "Market inflation expectation", False),
    ("DCOILWTICO", "Lagging Indicators", "WTI Crude Oil", "$/barrel", "Rising = inflationary pressure", False),
]

CATEGORY_ORDER = ["Leading Indicators", "Business Activity", "Coincident Indicators", "Lagging Indicators"]

HEADER_FILL = PatternFill("solid", fgColor="1F2A37")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)


def fetch_monthly(fred, fred_id):
    """Pull full history and resample to monthly. Quarterly series (GDP)
    get forward-filled across each month of their quarter; daily series
    (yields, WTI) get averaged to monthly; already-monthly series pass
    through as their month-start value."""
    raw = fred.get_series(fred_id, observation_start=HISTORY_START)
    raw = raw.dropna()
    if raw.empty:
        raise RuntimeError(f"FRED returned no data for {fred_id}")
    gaps = raw.index.to_series().diff().dropna()
    median_days = gaps.median().days if len(gaps) else 30
    if median_days <= 5:       # daily
        monthly = raw.resample("MS").mean()
    elif median_days <= 45:    # already monthly
        monthly = raw.resample("MS").last()
    else:                      # quarterly (GDP) -> backfill across the quarter
        monthly = raw.resample("MS").ffill()
        monthly = monthly.resample("MS").asfreq().ffill(limit=2)
    return monthly


def compute_yoy(monthly_series):
    return monthly_series.pct_change(12) * 100


def build_workbook(fred_api_key, output_path):
    fred = Fred(api_key=fred_api_key)

    pulled = {}
    for fred_id, category, name, unit, signal, needs_yoy in SERIES:
        print(f"Pulling {fred_id} ({name})...")
        monthly = fetch_monthly(fred, fred_id)
        if needs_yoy:
            monthly = compute_yoy(monthly).dropna()
        pulled[fred_id] = monthly

    all_dates = sorted(set().union(*[set(s.index) for s in pulled.values()]))

    wb = Workbook()
    wb.remove(wb.active)

    for category in CATEGORY_ORDER:
        ws = wb.create_sheet(category)
        header = ["FRED ID", "Display Name", "Unit", "Signal Logic"] + list(all_dates)
        ws.append(header)
        for cell in ws[1]:
            cell.font = HEADER_FONT
            cell.fill = HEADER_FILL
        for fred_id, cat, name, unit, signal, _ in SERIES:
            if cat != category:
                continue
            s = pulled[fred_id]
            row = [fred_id, name, unit, signal]
            for d in all_dates:
                v = s.get(d)
                row.append(None if v is None or pd.isna(v) else round(float(v), 4))
            ws.append(row)
        ws.column_dimensions["A"].width = 16
        ws.column_dimensions["B"].width = 28
        for i in range(len(all_dates)):
            ws.column_dimensions[get_column_letter(5 + i)].width = 10
        ws.freeze_panes = "E2"

    ws = wb.create_sheet("Dashboard", 0)
    ws.append(["MACRO DASHBOARD — 36 FRED Indicators"])
    ws.append([f"Source: FRED (St. Louis Fed) | Range: Jan-2019 to {date.today().strftime('%Y-%m')} | Pulled {date.today().strftime('%b %d, %Y')}"])
    ws.append([])
    ws.append(["Category", "FRED ID", "Display Name", "Unit", "Signal Logic",
               "Latest Month", "Latest Value", "Prior Value", "Change", "Status"])
    for cell in ws[4]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    for fred_id, category, name, unit, signal, _ in SERIES:
        s = pulled[fred_id].dropna()
        if len(s) < 2:
            continue
        latest_month, latest = s.index[-1], s.iloc[-1]
        prior = s.iloc[-2]
        change = latest - prior
        status = "Rising" if change > 0 else ("Falling" if change < 0 else "Flat")
        ws.append([category, fred_id, name, unit, signal, latest_month.strftime("%Y-%m"),
                   round(float(latest), 4), round(float(prior), 4), round(float(change), 4), status])

    ws.append([])
    ws.append(["Notes: Real GDP is quarterly, backfilled across each month of the quarter. "
               "WTI Crude and all daily-frequency series (Treasury yields, spreads, NFCI, breakevens) "
               "are averaged to monthly. CPI/Core CPI/Supercore/Core PCE/PPI are computed as 12-month "
               "% change from FRED's raw index-level series, not pulled as a native FRED %."])

    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["C"].width = 26

    wb.save(output_path)
    print(f"\nSaved: {output_path}")
    print(f"Series pulled: {len(SERIES)}  Date range: {all_dates[0].strftime('%Y-%m')} to {all_dates[-1].strftime('%Y-%m')}")


def main():
    ap = argparse.ArgumentParser(description="Pull the 36 FRED series macro_visualize.py needs.")
    ap.add_argument("-o", "--output", default="macro_dashboard.xlsx")
    ap.add_argument("--api-key", default=os.environ.get("FRED_API_KEY"))
    args = ap.parse_args()

    if not args.api_key:
        raise SystemExit(
            "No FRED_API_KEY found. Add it to your .env file, or pass --api-key YOUR_KEY."
        )
    build_workbook(args.api_key, args.output)


if __name__ == "__main__":
    main()
