"""
Portfolio — visualizer
=======================
Builds a standalone portfolio.html: same Bloomberg-terminal look as
terminal.html (black background, monospace, amber accent, one period
selector driving every row) — but this page's content is your full
holdings list, all positions, not just top/bottom movers.

Tracks LIVE prices — for free — pulled from a Google Sheet that uses
the GOOGLEFINANCE() formula.

WHY GOOGLE SHEETS
------------------
GOOGLEFINANCE() is Google's built-in, no-signup, no-API-key stock quote
function. The trick for pulling it into a script for free: put the
formula in a sheet, "publish" that sheet as a plain CSV file (Google
hosts and refreshes it for you, no login required to read it), then
have this script fetch that CSV like any other URL.

ACTUAL SHEET LAYOUT (as built) — column by column:
------------------------------------------------------
 A  Ticker
 B  (blank spacer column — ignored, kept because it's already there)
 C  Price          =GOOGLEFINANCE(A2,"price")
 D  1 Week
 E  1 Month
 F  3 Month
 G  6 Month
 H  1 Year         (NOT year-to-date — a trailing 12-month lookback)
 I  % Change Daily =GOOGLEFINANCE(A2,"changepct")
 J  T-1            your own number — prior-period target, whichever
                    basis (EPV or GV) you've already decided to use
 K  T              your own number — current-period target
 L  -15%           your own number — a separate modeled scenario, NOT
                    simply T*0.85 (your actual figures vary from -15%
                    to -46% off T depending on the position, so this
                    script reads it as its own independent input
                    rather than computing it)

D-H (the historical closes) use a small date window + INDEX so they
don't break on a weekend:
     D (1 week ago):  =INDEX(GOOGLEFINANCE($A2,"close",TODAY()-17,TODAY()-7),ROWS(GOOGLEFINANCE($A2,"close",TODAY()-17,TODAY()-7)),2)
     E (1 month ago): =INDEX(GOOGLEFINANCE($A2,"close",EDATE(TODAY(),-1)-10,EDATE(TODAY(),-1)),ROWS(GOOGLEFINANCE($A2,"close",EDATE(TODAY(),-1)-10,EDATE(TODAY(),-1))),2)
     F (3 months ago):=INDEX(GOOGLEFINANCE($A2,"close",EDATE(TODAY(),-3)-10,EDATE(TODAY(),-3)),ROWS(GOOGLEFINANCE($A2,"close",EDATE(TODAY(),-3)-10,EDATE(TODAY(),-3))),2)
     G (6 months ago):=INDEX(GOOGLEFINANCE($A2,"close",EDATE(TODAY(),-6)-10,EDATE(TODAY(),-6)),ROWS(GOOGLEFINANCE($A2,"close",EDATE(TODAY(),-6)-10,EDATE(TODAY(),-6))),2)
     H (1 year ago):  =INDEX(GOOGLEFINANCE($A2,"close",TODAY()-372,TODAY()-365),ROWS(GOOGLEFINANCE($A2,"close",TODAY()-372,TODAY()-365)),2)

J/K/L are plain numbers you type in yourself — your own valuation
work, not market data or a formula. There's no separate EPV/GV column
here: you decide which basis applies before typing the number in, so
the script just displays T-1/T/-15% generically without labeling
which method produced them. Leave a cell blank for anything not
valued yet; that position shows "—" until you fill it in. Re-run the
script any time after adding numbers — no code changes needed.

TODO (per Raman): none — the previous note here about STRL's T-1/T/-15%
being pre-split/stale was wrong. Those are the real, intended values.

5. If a ticker errors anywhere, wrap it with its exchange inside the
   formula, e.g. GOOGLEFINANCE("NYSE:AXP",...) — most US large-caps
   resolve fine bare, but some ADRs and small-caps need it.
6. Get the CSV export URL — two ways, pick one:

   RECOMMENDED — direct export link, no publishing step needed:
     https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/export?format=csv&gid=0
   Just needs Share -> General access -> "Anyone with the link" set to
   Viewer. Find YOUR_SHEET_ID in the sheet's normal /d/.../edit URL —
   it's the long ID between /d/ and /edit. gid=0 is the first tab;
   check the URL bar for a different gid if you're using another tab.

   ALTERNATIVE — File -> Share -> Publish to web, format "Comma-
   separated values (.csv)", Publish. Gives a different-looking URL
   ending in output=csv. Slightly more setup (an extra publish step,
   and the sheet's contents become more discoverable), but works even
   if you'd rather not change the share-link permission above.

   Either way, the URL you end up with needs to return raw CSV text
   when fetched — not an HTML page. The plain /edit link never works
   for this; it requires a signed-in browser session a script doesn't
   have.
7. Copy that URL into SHEET_CSV_URL below.

Google serves both URL types with only a few minutes of lag from the
sheet's real state, so running this script every couple hours (same
cadence as the macro dashboard) is more than fast enough to catch
fresh prices — no polling tricks needed, and nothing here requires a
paid API or a login.

HOLDINGS TAB (optional) — edit shares/cost/cash in the sheet, not code
------------------------------------------------------------------------
By default, shares/cost basis/cash are hardcoded in this file (the
HOLDINGS list and CASH_VALUE below). To edit them from the sheet
instead, add a new TAB (not a column — an actual second sheet tab)
called exactly "Holdings" in the same spreadsheet:

  Row 1 (header): Ticker | Shares | Cost Basis | Cash Value
  Row 2 onward:   MEDP   | 11857  | 249.55     | (blank)
                  SEI    | 83015  | 19.92      | (blank)
                  ...
                  CASH   | (blank)| (blank)    | 5475000

"CASH" is a special row — only its Cash Value column matters, Shares/
Cost stay blank. Every other row is a real position — Shares and Cost
Basis both required, Cash Value stays blank.

To buy/sell something or update your cash balance going forward: edit
this tab, not this file. Add or remove a whole row to add/drop a
position entirely.

For the live automated version, this tab needs its OWN "Publish to
web" CSV link (different tab = different URL, same as everything else
in this project) — publish it the same way as the main sheet, then
paste that URL into HOLDINGS_CSV_URL below.

If no "Holdings" tab exists, or that URL is still the placeholder,
this script quietly falls back to the hardcoded HOLDINGS/CASH_VALUE
below — nothing breaks if you haven't set this up yet.

USAGE
-----
python portfolio_visualize.py -o portfolio.html
"""

import argparse
import csv
import io
import json
import urllib.request
from datetime import datetime

# --- Fill this in after step 5 above ---
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRRTQgiz5Lz7_dT5EXDOFUJK4rWQv2xNPghffjQATAoj1rCajndsUPBkAVxNQVgagULrJ0yfzsp4zY6/pub?output=csv"

# Holdings that don't change tick-to-tick: shares and cost basis. Update
# this list only when you actually buy or sell — prices come live from
# the sheet, so weight/value/return recompute automatically every run.
# --- Fallback defaults, used only for --demo or if no Holdings tab is
# found. Once the "Holdings" sheet tab exists (see load_holdings_xlsx /
# load_holdings_csv below and the docstring), these get overwritten at
# runtime — you edit the sheet, not this file, from that point on.
# (symbol, shares, cost_basis)
HOLDINGS = [
    ("MEDP", 11857, 249.55),
    ("SEI", 83015, 19.92),
    ("SHEL", 52574, 70.26),
    ("CROX", 39278, 102.80),
    ("MHO", 34232, 139.22),
    ("RS", 10006, 255.27),
    ("FTNT", 73376, 66.27),
    ("MLI", 64280, 21.60),
    ("NVDA", 30113, 116.17),
    ("IAU", 58900, 46.51),
    ("META", 10162, 261.22),
    ("MSFT", 13754, 370.20),
    ("APH", 78972, 73.83),
    ("AXP", 16549, 154.60),
    ("AMAT", 13151, 130.84),
    ("STRL", 13058, 354.30),
    ("AVGO", 14100, 173.85),
    ("SNEX", 98061, 35.08),
    ("AZO", 1083, 3162.27),
    ("IBIT", 126192, 47.96),
    ("ONEW", 296210, 21.32),
    ("ASO", 87396, 46.09),
]
CASH_VALUE = 5475000.00

# Where the live Holdings tab is published (separate from SHEET_CSV_URL
# above, since each Google Sheets tab needs its own "Publish to web"
# link — different tab, different URL). See the docstring for the exact
# tab layout to set up.
HOLDINGS_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRRTQgiz5Lz7_dT5EXDOFUJK4rWQv2xNPghffjQATAoj1rCajndsUPBkAVxNQVgagULrJ0yfzsp4zY6/pub?gid=1480920284&single=true&output=csv"


def _parse_num(v):
    """Handles the real Holdings export's formatting: numbers as plain
    values, or as strings with commas/$ signs ("13,151", "$249.55"), and
    non-numeric placeholders like "Add lot" (PSQ's row, no real
    position) which should be skipped, not crash the parse."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace("$", "").strip()
    try:
        return float(s)
    except ValueError:
        return None  # e.g. "Add lot"


def load_holdings_xlsx(path, sheet_name="Holdings"):
    """Overwrites the module-level HOLDINGS/CASH_VALUE from a "Holdings"
    tab in the same workbook, if one exists. This reads your actual
    brokerage-style export layout — Symbol(A), Shares(G), Cost(H),
    Value(N) — not a simplified format; skips rows with no real
    position (PSQ's "Add lot") and the TOTAL summary row. Silently
    leaves the hardcoded defaults in place if the tab isn't there yet."""
    global HOLDINGS, CASH_VALUE
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        return False
    ws = wb[sheet_name]
    new_holdings = []
    new_cash = CASH_VALUE
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        ticker = str(row[0]).strip().upper()
        if ticker in ("TOTAL",):
            continue
        if ticker == "CASH":
            cash_val = _parse_num(row[13]) if len(row) > 13 else None  # column N
            if cash_val is not None:
                new_cash = cash_val
            continue
        shares = _parse_num(row[6]) if len(row) > 6 else None   # column G
        cost = _parse_num(row[7]) if len(row) > 7 else None     # column H
        if shares is None or cost is None:
            continue  # e.g. PSQ's "Add lot" — no real position to track
        new_holdings.append((ticker, shares, cost))
    if new_holdings:
        HOLDINGS = new_holdings
        CASH_VALUE = new_cash
        return True
    return False


def load_holdings_csv(csv_url):
    """Same as load_holdings_xlsx but from the Holdings tab's own
    published CSV URL. Silently leaves defaults in place if the URL
    isn't set up yet."""
    global HOLDINGS, CASH_VALUE
    if not csv_url or "PASTE_YOUR" in csv_url:
        return False
    with urllib.request.urlopen(csv_url, timeout=20) as resp:
        text = resp.read().decode("utf-8")
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    new_holdings = []
    new_cash = CASH_VALUE
    for row in rows[1:]:
        if not row or not row[0].strip():
            continue
        ticker = row[0].strip().upper()
        if ticker in ("TOTAL",):
            continue
        if ticker == "CASH":
            cash_val = _parse_num(row[13]) if len(row) > 13 else None  # column N
            if cash_val is not None:
                new_cash = cash_val
            continue
        shares = _parse_num(row[6]) if len(row) > 6 else None   # column G
        cost = _parse_num(row[7]) if len(row) > 7 else None     # column H
        if shares is None or cost is None:
            continue  # e.g. PSQ's "Add lot" — no real position to track
        new_holdings.append((ticker, shares, cost))
    if new_holdings:
        HOLDINGS = new_holdings
        CASH_VALUE = new_cash
        return True
    return False

# Price targets used to live here as a hardcoded dict — they've moved to
# columns I (EPV) and J (GV) in the Google Sheet instead, so you can type
# new values in as you finish valuation work without touching this file.
# See the docstring above for the exact columns.


def read_xlsx_price_data(path):
    """Reads the ACTUAL sheet layout: Ticker, blank, Price, 1W, 1M, 3M,
    6M, 1Y, Daily%, T-1, T, -15% — read from a local .xlsx export, no
    header row assumed. Column I (Daily%) is a direct percentage from
    GOOGLEFINANCE's changepct; columns J/K/L (T-1/T/-15%) are numbers
    you type in yourself — see this script's docstring for the exact
    column-by-column layout."""
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    data = {}
    for row in ws.iter_rows(values_only=True):
        if not row or not row[0]:
            continue
        ticker = str(row[0]).strip().upper()
        if ticker == "TICKER":  # tolerate an optional header row
            continue

        def cell(i):
            return row[i] if i < len(row) else None

        data[ticker] = {
            # row[1] is the blank spacer column — intentionally skipped
            "price": cell(2), "1w": cell(3), "1m": cell(4),
            "3m": cell(5), "6m": cell(6), "1y": cell(7), "1d_pct": cell(8),
            "t1": cell(9), "t": cell(10), "m15": cell(11),
            "ytd_close": cell(15),  # column P, shared with porter_visualize.py
        }
    return data


def fetch_price_data(csv_url):
    if "PASTE_YOUR" in csv_url:
        raise SystemExit(
            "SHEET_CSV_URL isn't set yet — see the setup steps in the top "
            "of this script (publish your Google Sheet as CSV first)."
        )
    with urllib.request.urlopen(csv_url, timeout=20) as resp:
        text = resp.read().decode("utf-8")
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    data = {}
    # Columns: A Ticker, B (blank), C Price, D 1W, E 1M, F 3M, G 6M, H 1Y,
    #          I Daily%, J T-1, K T, L -15%, ... P YTD-close (shared w/ porter_visualize.py)
    for row in rows[1:]:  # skip header
        if len(row) < 2 or not row[0].strip():
            continue
        ticker = row[0].strip().upper()

        def cell(i):
            if i >= len(row) or row[i] == "":
                return None
            try:
                return float(row[i].replace(",", ""))
            except (ValueError, AttributeError):
                return None  # blank / #N/A / not published yet

        data[ticker] = {
            "price": cell(2), "1w": cell(3), "1m": cell(4),
            "3m": cell(5), "6m": cell(6), "1y": cell(7), "1d_pct": cell(8),
            "t1": cell(9), "t": cell(10), "m15": cell(11),
            "ytd_close": cell(15),
        }
    return data


def demo_price_data():
    """Synthetic prices near cost basis, for testing the render without a
    live sheet wired up yet."""
    import random
    rng = random.Random(7)
    data = {}
    for sym, _, cost in HOLDINGS:
        price = round(cost * (1 + rng.uniform(-0.3, 1.2)), 2)
        data[sym] = {
            "price": price,
            "1w": round(price * (1 + rng.uniform(-0.05, 0.05)), 2),
            "1m": round(price * (1 + rng.uniform(-0.1, 0.1)), 2),
            "3m": round(price * (1 + rng.uniform(-0.2, 0.2)), 2),
            "6m": round(price * (1 + rng.uniform(-0.3, 0.3)), 2),
            "1y": round(price * (1 + rng.uniform(-0.25, 0.25)), 2),
            "1d_pct": round(rng.uniform(-4, 4), 2),
            "t1": None, "t": None, "m15": None,
            "ytd_close": round(price * (1 + rng.uniform(-0.2, 0.2)), 2),
        }
    # give the demo a populated example so the render is testable
    if "STRL" in data:
        data["STRL"]["t1"] = 308.54
        data["STRL"]["t"] = 502.45
        data["STRL"]["m15"] = 427.08
    # SPY isn't a holding — read separately as a benchmark, not part of HOLDINGS
    data["SPY"] = {
        "price": 680.0, "1w": 675.0, "1m": 660.0, "3m": 640.0,
        "6m": 600.0, "1y": 560.0, "1d_pct": 0.3, "ytd_close": 610.0,
        "t1": None, "t": None, "m15": None,
    }
    return data


PERIODS = ["1w", "1m", "3m", "6m", "1y"]
PERIOD_LABELS = {"1d": "Daily", "1w": "1W", "3m": "3M", "6m": "6M", "1y": "1Y", "total": "Total"}


def pct_change(current, prior):
    if current is None or prior is None or prior == 0:
        return None
    return (current - prior) / prior * 100


def build_rows(data):
    enriched = []
    for symbol, shares, cost in HOLDINGS:
        d = data.get(symbol, {})
        price = d.get("price")
        row = {"symbol": symbol, "shares": shares, "cost": cost, "price": price}
        if price is None:
            row["value"] = None
            row["total_pct"] = None
        else:
            row["value"] = shares * price
            row["total_pct"] = pct_change(price, cost)
        for p in PERIODS:
            row[p] = pct_change(price, d.get(p)) if price is not None else None
        # 1d comes as a ready-made % from GOOGLEFINANCE's changepct, not a
        # price to diff — pass it through as-is rather than computing it.
        row["1d"] = d.get("1d_pct") if price is not None else None
        # T-1 / T / -15% are all independent numbers you typed in — not
        # derived from each other (your own figures range from -15% to
        # -46% off T, so -15% is never computed as T*0.85 here).
        row["t1"] = d.get("t1")
        row["t"] = d.get("t")
        row["m15"] = d.get("m15")
        row["ytd"] = pct_change(price, d.get("ytd_close")) if price is not None else None
        enriched.append(row)
    return enriched


def _valued(price, target_price):
    return {"price": target_price, "upside": (target_price - price) / price * 100}


def build_positions_json(rows, total_value):
    positions = []
    for r in sorted(rows, key=lambda r: (r["value"] is None, -(r["value"] or 0))):
        weight = None if r["value"] is None else r["value"] / total_value * 100
        # No EPV/GV distinction in the sheet — you decide which basis
        # applies before typing the number in, so this just reads
        # whatever's in T-1/T/-15% directly, generically.
        target = None
        if r["price"] is not None and r.get("t") is not None:
            target = {
                "t1": _valued(r["price"], r["t1"]) if r.get("t1") is not None else None,
                "t": _valued(r["price"], r["t"]),
                "m15": _valued(r["price"], r["m15"]) if r.get("m15") is not None else None,
            }
        positions.append({
            "symbol": r["symbol"], "price": r["price"], "value": r["value"], "weight": weight,
            "1d": r["1d"], "1w": r["1w"], "1m": r["1m"], "3m": r["3m"], "6m": r["6m"],
            "1y": r["1y"], "ytd": r["ytd"], "total": r["total_pct"], "target": target,
        })
    return positions


def build_spy_json(data):
    """SPY isn't a holding — no shares/cost basis — just a benchmark row
    read straight from the sheet if it's there. Add a "SPY" row to the
    same sheet (same columns as any ticker: price, 1W/1M/3M/6M/1Y/YTD
    closes) to enable this; without it, the comparison just doesn't
    render rather than showing fake numbers."""
    d = data.get("SPY")
    if not d or d.get("price") is None:
        return None
    price = d["price"]
    return {
        "price": price,
        "1d": d.get("1d_pct"),
        "1w": pct_change(price, d.get("1w")),
        "1m": pct_change(price, d.get("1m")),
        "3m": pct_change(price, d.get("3m")),
        "6m": pct_change(price, d.get("6m")),
        "1y": pct_change(price, d.get("1y")),
        "ytd": pct_change(price, d.get("ytd_close")),
        # SPY has no cost basis, so "Total" (which for positions means
        # vs. cost) doesn't apply — left unset rather than faked.
        "total": None,
    }


HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>KWP Terminal &mdash; Portfolio</title>
<style>
  :root {
    --bg: #000000;
    --panel: #0a0a0a;
    --line: #262626;
    --amber: #ffb000;
    --white: #d8d8d8;
    --green: #00d964;
    --red: #ff3b3b;
    --mono: "Menlo", "Consolas", "SFMono-Regular", monospace;
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--bg); color: var(--white); font-family: var(--mono); }

  header { display: flex; justify-content: space-between; align-items: center; padding: 10px 16px; border-bottom: 1px solid var(--line); flex-wrap: wrap; gap: 8px; }
  .brand { color: var(--amber); font-size: 14px; letter-spacing: 2px; }
  nav.fkeys { display: flex; gap: 2px; }
  nav.fkeys a {
    color: var(--white); text-decoration: none; font-size: 11px; padding: 5px 10px;
    border: 1px solid var(--line); background: var(--panel);
  }
  nav.fkeys a.active { color: #000; background: var(--amber); border-color: var(--amber); }
  .as-of { font-size: 10px; color: #666; }

  .wrap { padding: 14px 16px 40px; max-width: 1200px; margin: 0 auto; }
  .total-strip { display: flex; gap: 30px; align-items: baseline; margin: 16px 0; padding-bottom: 12px; border-bottom: 1px solid var(--line); flex-wrap: wrap; }
  .change-period-bar { display: flex; gap: 2px; align-self: center; margin-left: auto; }
  .change-period-bar button {
    font-family: var(--mono); font-size: 10px; color: var(--white); background: var(--panel);
    border: 1px solid var(--line); padding: 4px 10px; cursor: pointer;
  }
  .change-period-bar button.active { color: #000; background: var(--amber); border-color: var(--amber); }
  .total-label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 1px; }
  .total-value { font-size: 26px; color: var(--amber); }

  .panel { background: var(--panel); border: 1px solid var(--line); }
  .panel h3 {
    margin: 0; padding: 6px 10px; font-size: 10.5px; color: #000; background: var(--amber);
    text-transform: uppercase; letter-spacing: 1px; font-weight: normal;
  }
  .p-scroll { overflow-x: auto; }
  .p-row {
    display: grid; grid-template-columns: 65px 65px 80px 65px 65px 65px 65px 65px 65px 70px 100px 90px 100px 90px;
    gap: 8px; padding: 7px 12px; font-size: 12px; border-top: 1px solid var(--line);
    align-items: center; min-width: 1320px;
  }
  .p-row:first-child { border-top: none; }
  .p-row.p-head { color: #888; font-size: 9.5px; text-transform: uppercase; letter-spacing: 0.5px; background: #111; }
  .sortable { cursor: pointer; user-select: none; }
  .sortable:hover { color: var(--amber); }
  .p-symbol { color: var(--amber); }
  .p-num { text-align: right; }
  .p-t-main { font-weight: bold; font-size: 14px; }
  .t-up { color: var(--green); }
  .t-down { color: var(--red); }
  .t-na { color: #555; font-size: 11px; }
  .t-error { color: var(--red); font-size: 11px; }

  .note { font-size: 11px; color: #666; margin-top: 12px; }
  .warn { font-size: 11px; color: var(--red); margin-top: 6px; }
  footer { text-align: center; font-size: 10px; color: #444; margin-top: 24px; }
</style>
</head>
<body>
  <header>
    <div class="brand">KWP TERMINAL</div>
    <nav class="fkeys">
      <a href="index.html">[F1] MACRO</a>
      <a href="portfolio.html" class="active">[F2] PORTFOLIO</a>
      <a href="terminal.html">[F3] TERMINAL</a>
      <a href="porter.html">[F4] PORTER</a>
    </nav>
    <div class="as-of">__GENDATE__</div>
  </header>

  <div class="wrap">
    <div class="total-strip">
      <div><div class="total-label">Total Value</div><div class="total-value">__TOTAL__</div></div>
      <div><div class="total-label" id="change-period-label">Daily $ Change</div><div class="total-value" id="total-dollar-change"></div></div>
      <div><div class="total-label" id="change-period-label-2">Daily % Change</div><div class="total-value" id="total-pct-change"></div></div>
      <div><div class="total-label" id="spy-period-label">SPY Daily %</div><div class="total-value" id="spy-pct-change"></div></div>
      <div class="change-period-bar" id="change-period-bar">
        <button data-period="1d" class="active">Daily</button>
        <button data-period="1w">1W</button>
        <button data-period="1m">1M</button>
        <button data-period="3m">3M</button>
        <button data-period="6m">6M</button>
        <button data-period="ytd">YTD</button>
        <button data-period="1y">1Y</button>
        <button data-period="total">Total</button>
      </div>
    </div>

    <div class="panel">
      <h3>Holdings</h3>
      <div class="p-scroll"><div id="holdings"></div></div>
    </div>

    <div class="note">Total $/% Change is vs. cost basis (same as the Total column). T-1/T/-15% all come from columns J-L in your sheet — your own numbers, whichever basis (EPV or GV) you've already applied. Daily needs the changepct sheet column (see this script's docstring).</div>
    __MISSING_WARNING__

    <footer>Prices via GOOGLEFINANCE, published from your own Google Sheet &middot; generated __GENDATE__</footer>
  </div>

<script>
const POSITIONS = __POSITIONS__;
const SPY = __SPY__;
const CASH_VALUE = __CASH_VALUE__;
const TOTAL_VALUE = __TOTAL_VALUE_NUM__;

const SORT_COLUMNS = [
  { key: "symbol", label: "Symbol", type: "text", numeric: false },
  { key: "weight", label: "Weight", type: "num" },
  { key: "price", label: "Price", type: "num" },
  { key: "1d", label: "Daily", type: "num" },
  { key: "1w", label: "1W", type: "num" },
  { key: "3m", label: "3M", type: "num" },
  { key: "6m", label: "6M", type: "num" },
  { key: "ytd", label: "YTD", type: "num" },
  { key: "1y", label: "1Y", type: "num" },
  { key: "total", label: "Total", type: "num" },
  { key: "value", label: "Value", type: "num" },
  { key: "t1", label: "T-1", type: "target" },
  { key: "t", label: "T", type: "target" },
  { key: "m15", label: "-15%", type: "target" },
];

let sortKey = "weight";
let sortDir = "desc"; // default matches the existing weight-descending order

function sortValue(p, col) {
  if (col.type === "text") return p.symbol;
  if (col.type === "target") return p.target ? p.target[col.key]?.upside : null;
  return p[col.key];
}

function pctHtml(pct) {
  if (pct === null || pct === undefined) return '<span class="t-na">N/A</span>';
  const cls = pct >= 0 ? "t-up" : "t-down";
  const sign = pct >= 0 ? "+" : "";
  return `<span class="${cls}">${sign}${pct.toFixed(1)}%</span>`;
}

function valCell(v, emphasize) {
  if (!v) return `<div class="p-num${emphasize ? ' p-t-main' : ''}">&mdash;</div>`;
  const cls = v.upside >= 0 ? "t-up" : "t-down";
  const sign = v.upside >= 0 ? "+" : "";
  return `<div class="p-num${emphasize ? ' p-t-main' : ''}">$${v.price.toFixed(2)} <span class="${cls}">(${sign}${v.upside.toFixed(1)}%)</span></div>`;
}

function renderHoldings() {
  const el = document.getElementById("holdings");

  const head = `<div class="p-row p-head">` + SORT_COLUMNS.map(col => {
    const active = col.key === sortKey;
    const arrow = active ? (sortDir === "desc" ? " \u25BC" : " \u25B2") : "";
    const cls = col.type === "text" ? "sortable" : "p-num sortable";
    return `<div class="${cls}" data-sort="${col.key}">${col.label}${arrow}</div>`;
  }).join("") + `</div>`;

  // Sort a copy — priced positions only; unpriced ("no quote") rows
  // always sort last regardless of column, since there's nothing to
  // compare. Nulls within priced positions also sort last.
  const sorted = [...POSITIONS].sort((a, b) => {
    if (a.price === null && b.price === null) return 0;
    if (a.price === null) return 1;
    if (b.price === null) return -1;
    const col = SORT_COLUMNS.find(c => c.key === sortKey);
    const va = sortValue(a, col), vb = sortValue(b, col);
    if (va === null || va === undefined) return 1;
    if (vb === null || vb === undefined) return -1;
    if (col.type === "text") {
      return sortDir === "asc" ? va.localeCompare(vb) : vb.localeCompare(va);
    }
    return sortDir === "asc" ? va - vb : vb - va;
  });

  const rows = sorted.map(p => {
    if (p.price === null) {
      return `<div class="p-row"><div class="p-symbol">${p.symbol}</div>
        <div class="p-num">&mdash;</div><div class="p-num t-error">no quote</div>
        ${'<div class="p-num">&mdash;</div>'.repeat(7)}
        <div class="p-num">&mdash;</div>
        ${'<div class="p-num">&mdash;</div>'.repeat(3)}</div>`;
    }
    const t = p.target;
    return `<div class="p-row"><div class="p-symbol">${p.symbol}</div>
      <div class="p-num">${p.weight.toFixed(2)}%</div>
      <div class="p-num">$${p.price.toFixed(2)}</div>
      <div class="p-num">${pctHtml(p["1d"])}</div>
      <div class="p-num">${pctHtml(p["1w"])}</div>
      <div class="p-num">${pctHtml(p["3m"])}</div>
      <div class="p-num">${pctHtml(p["6m"])}</div>
      <div class="p-num">${pctHtml(p["ytd"])}</div>
      <div class="p-num">${pctHtml(p["1y"])}</div>
      <div class="p-num">${pctHtml(p["total"])}</div>
      <div class="p-num">$${p.value.toLocaleString(undefined, {maximumFractionDigits: 0})}</div>
      ${valCell(t && t.t1, false)}
      ${valCell(t && t.t, true)}
      ${valCell(t && t.m15, false)}</div>`;
  }).join("");

  const cashWeight = CASH_VALUE / TOTAL_VALUE * 100;
  const cashRow = `<div class="p-row"><div class="p-symbol">CASH</div>
    <div class="p-num">${cashWeight.toFixed(2)}%</div>
    ${'<div class="p-num">&mdash;</div>'.repeat(8)}
    <div class="p-num">$${CASH_VALUE.toLocaleString(undefined, {maximumFractionDigits: 0})}</div>
    ${'<div class="p-num">&mdash;</div>'.repeat(3)}</div>`;

  el.innerHTML = head + rows + cashRow;
}

const CHANGE_PERIOD_LABEL = {"1d": "Daily", "1w": "1W", "1m": "1M", "3m": "3M", "6m": "6M", "ytd": "YTD", "1y": "1Y", "total": "Total"};
let changePeriod = "1d";

function renderTotalChange() {
  const label = CHANGE_PERIOD_LABEL[changePeriod];
  document.getElementById("change-period-label").textContent = `${label} $ Change`;
  document.getElementById("change-period-label-2").textContent = `${label} % Change`;
  document.getElementById("spy-period-label").textContent = `SPY ${label} %`;

  // Sum each position's dollar change for the selected period (cash
  // contributes 0 — no cost basis or price to move against). "Total"
  // uses the same vs.-cost-basis math as the Total column; every other
  // period derives the prior value from that period's % change, same
  // trick used elsewhere: pct = (now-then)/then, so then = now/(1+pct/100).
  let totalDollarChange = 0;
  let anyData = false;
  POSITIONS.forEach(p => {
    const pct = changePeriod === "total" ? p.total : p[changePeriod];
    if (p.value !== null && pct !== null && pct !== undefined) {
      const priorValue = p.value / (1 + pct / 100);
      totalDollarChange += p.value - priorValue;
      anyData = true;
    }
  });

  const dollarEl = document.getElementById("total-dollar-change");
  const pctEl = document.getElementById("total-pct-change");
  if (!anyData) {
    dollarEl.innerHTML = '<span class="t-na">N/A</span>';
    pctEl.innerHTML = '<span class="t-na">N/A</span>';
  } else {
    const priorTotal = TOTAL_VALUE - totalDollarChange;
    const totalPctChange = priorTotal !== 0 ? (totalDollarChange / priorTotal) * 100 : null;
    const cls = totalDollarChange >= 0 ? "t-up" : "t-down";
    const sign = totalDollarChange >= 0 ? "+" : "";
    dollarEl.innerHTML = `<span class="${cls}">${sign}$${Math.abs(totalDollarChange).toLocaleString(undefined, {maximumFractionDigits: 0})}</span>`;
    pctEl.innerHTML = totalPctChange === null ? '<span class="t-na">N/A</span>' : `<span class="${cls}">${sign}${totalPctChange.toFixed(2)}%</span>`;
  }

  // SPY comparison — separate from the block above so a missing SPY row
  // (or a period with no SPY data, e.g. Total, which SPY has no cost
  // basis for) never blocks your own portfolio numbers from showing.
  const spyEl = document.getElementById("spy-pct-change");
  const spyPct = SPY ? SPY[changePeriod] : null;
  if (spyPct === null || spyPct === undefined) {
    spyEl.innerHTML = '<span class="t-na">N/A</span>';
  } else {
    const cls = spyPct >= 0 ? "t-up" : "t-down";
    const sign = spyPct >= 0 ? "+" : "";
    spyEl.innerHTML = `<span class="${cls}">${sign}${spyPct.toFixed(2)}%</span>`;
  }
}

renderHoldings();
renderTotalChange();

document.getElementById("change-period-bar").addEventListener("click", e => {
  const btn = e.target.closest("button[data-period]");
  if (!btn) return;
  changePeriod = btn.dataset.period;
  document.querySelectorAll("#change-period-bar button").forEach(b => b.classList.toggle("active", b === btn));
  renderTotalChange();
});

document.getElementById("holdings").addEventListener("click", e => {
  const cell = e.target.closest(".sortable");
  if (!cell) return;
  const key = cell.dataset.sort;
  if (key === sortKey) {
    sortDir = sortDir === "desc" ? "asc" : "desc";
  } else {
    sortKey = key;
    // Numeric columns default to descending (biggest first); Symbol
    // defaults to ascending (A-Z) since that's the natural reading order.
    sortDir = key === "symbol" ? "asc" : "desc";
  }
  renderHoldings();
});
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser(description="Build the live portfolio.html page.")
    ap.add_argument("-o", "--output", default="portfolio.html")
    ap.add_argument("--sheet-csv-url", default=SHEET_CSV_URL)
    ap.add_argument("--holdings-csv-url", default=HOLDINGS_CSV_URL, help="separate published CSV for the Holdings tab")
    ap.add_argument("--xlsx", default=None, help="read prices from a local .xlsx export instead of the live sheet URL")
    ap.add_argument("--demo", action="store_true", help="use synthetic prices, no sheet needed")
    args = ap.parse_args()

    if args.demo:
        prices = demo_price_data()
    elif args.xlsx:
        prices = read_xlsx_price_data(args.xlsx)
        # Same file may also have a "Holdings" tab — if so, shares/cost/
        # cash come from there instead of the hardcoded defaults.
        load_holdings_xlsx(args.xlsx)
    else:
        prices = fetch_price_data(args.sheet_csv_url)
        load_holdings_csv(args.holdings_csv_url)
    rows = build_rows(prices)

    priced_total = sum(r["value"] for r in rows if r["value"] is not None)
    total_value = priced_total + CASH_VALUE
    positions = build_positions_json(rows, total_value)
    spy = build_spy_json(prices)
    missing = [r["symbol"] for r in rows if r["price"] is None]

    html = HTML_TEMPLATE.replace("__TOTAL__", f"${total_value:,.0f}")
    html = html.replace("__TOTAL_VALUE_NUM__", json.dumps(total_value))
    html = html.replace("__CASH_VALUE__", json.dumps(CASH_VALUE))
    html = html.replace("__POSITIONS__", json.dumps(positions))
    html = html.replace("__SPY__", json.dumps(spy))
    html = html.replace("__GENDATE__", datetime.today().strftime("%Y-%m-%d %H:%M"))
    warning = ""
    if missing:
        warning = f'<div class="warn">No quote returned for: {", ".join(missing)} — check the ticker/formula in your sheet.</div>'
    html = html.replace("__MISSING_WARNING__", warning)

    with open(args.output, "w") as f:
        f.write(html)
    print(f"Wrote {args.output} ({len(rows)} positions, {len(missing)} missing quotes)")


if __name__ == "__main__":
    main()
