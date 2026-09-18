"""
Porter — visualizer
=====================
A research/fundamentals view of your holdings: Symbol, Price, Change,
P/E, D/E, EBIT Margin, Industry, Earnings Date, and performance across
5D / 1W / 1M / 6M / YTD / 1Y — sortable by any column, same click-to-sort
pattern as portfolio.html. Built for your boss to look at, separate
from Macro / Portfolio / Terminal.

WHAT'S HERE
-------------
All 10 columns from your original reference screenshot are covered.
Six are live from GOOGLEFINANCE (Price, Change, Change %, P/E, and the
Perf columns). Four are manually-entered numbers you type into the
sheet yourself, since none of them are available from GOOGLEFINANCE at
any price tier: Debt to Equity, EBIT Margin, Industry/Sub Class, and
Upcoming Announce Date. Same pattern as T-1/T/-15% on portfolio.html —
blank cells just show N/A rather than breaking.

SHEET SETUP — 8 new columns beyond what portfolio_visualize.py uses
----------------------------------------------------------------------
Same "KWP" sheet, same tickers in column A. Add these:
  M  P/E (TTM)        =GOOGLEFINANCE(A2,"pe")
  N  Change ($)       =GOOGLEFINANCE(A2,"change")
  O  5D ago close      =INDEX(GOOGLEFINANCE($A2,"close",TODAY()-11,TODAY()-5),ROWS(GOOGLEFINANCE($A2,"close",TODAY()-11,TODAY()-5)),2)
  P  YTD baseline close =INDEX(GOOGLEFINANCE($A2,"close",DATE(YEAR(TODAY()),1,1),DATE(YEAR(TODAY()),1,10)),2,2)
  Q  Debt to Equity    typed in manually — not available from GOOGLEFINANCE
  R  EBIT Margin       typed in manually — not available from GOOGLEFINANCE
  S  Industry/Sub Class typed in manually — not available from GOOGLEFINANCE
  T  Upcoming Announce Date typed in manually — not available from GOOGLEFINANCE

Cells containing the literal text "#N/A" (e.g. P/E for ETFs, which
don't have one) are treated as missing data, not as a value — see
_num() below.

Everything else this page needs (Price, Daily %, 1W, 1M, 6M, 1Y) is
already in columns C, I, D, E, G, H from the existing sheet setup —
see portfolio_visualize.py's docstring for those.

Note on 5D vs 1W: these will often be very close to each other (5
trading days ~ 1 calendar week), but they're genuinely different
lookback windows, not the same number relabeled.

USAGE
-----
python porter_visualize.py -o porter.html
python porter_visualize.py --xlsx KWP.xlsx -o porter.html
python porter_visualize.py --demo -o porter.html
"""

import argparse
import csv
import io
import json
import urllib.request
from datetime import datetime

import portfolio_visualize as pv

# Single source of truth: reuse portfolio_visualize.py's URL rather than
# maintaining a second copy that could drift out of sync with it.
SHEET_CSV_URL = pv.SHEET_CSV_URL


def get_tickers():
    """Derived from portfolio_visualize.py's HOLDINGS at call time (not
    a static list here) so this page always shows whatever you actually
    hold — including anything loaded from the Holdings sheet tab — never
    a second, separately-maintained ticker list that could drift out of
    sync with Portfolio's."""
    return [h[0] for h in pv.HOLDINGS]


def _num(v):
    """Safely coerce a cell value to a float, treating '#N/A' and other
    non-numeric text as missing data rather than crashing or passing a
    literal error string through to the page."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ""))
    except ValueError:
        return None  # e.g. the literal text "#N/A"


def read_xlsx_data(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    data = {}
    for row in ws.iter_rows(values_only=True):
        if not row or not row[0]:
            continue
        ticker = str(row[0]).strip().upper()
        if ticker == "TICKER":
            continue

        def cell(i):
            return row[i] if i < len(row) else None

        data[ticker] = {
            "price": _num(cell(2)), "1w": _num(cell(3)), "1m": _num(cell(4)),
            "6m": _num(cell(6)), "1y": _num(cell(7)), "change_pct": _num(cell(8)),
            "pe": _num(cell(12)), "change_dollar": _num(cell(13)),
            "5d": _num(cell(14)), "ytd": _num(cell(15)), "debt_equity": _num(cell(16)),
            "ebit_margin": _num(cell(17)),
            "industry": cell(18) if cell(18) else None,
            "earnings_date": cell(19).strftime("%Y-%m-%d") if hasattr(cell(19), "strftime") else cell(19),
        }
    return data


def fetch_csv_data(csv_url):
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
    for row in rows[1:]:
        if len(row) < 2 or not row[0].strip():
            continue
        ticker = row[0].strip().upper()

        def cell(i):
            return row[i] if i < len(row) and row[i] != "" else None

        data[ticker] = {
            "price": _num(cell(2)), "1w": _num(cell(3)), "1m": _num(cell(4)),
            "6m": _num(cell(6)), "1y": _num(cell(7)), "change_pct": _num(cell(8)),
            "pe": _num(cell(12)), "change_dollar": _num(cell(13)),
            "5d": _num(cell(14)), "ytd": _num(cell(15)), "debt_equity": _num(cell(16)),
            "ebit_margin": _num(cell(17)), "industry": cell(18), "earnings_date": cell(19),
        }
    return data


def demo_data():
    import random
    rng = random.Random(11)
    industries = ["Technology / Software", "Healthcare / Devices", "Energy / Equipment",
                  "Financials / Capital Markets", "Industrials / Machinery"]
    data = {}
    for sym in get_tickers():
        price = round(rng.uniform(20, 700), 2)
        data[sym] = {
            "price": price,
            "1w": round(price * (1 + rng.uniform(-0.05, 0.05)), 2),
            "1m": round(price * (1 + rng.uniform(-0.1, 0.1)), 2),
            "6m": round(price * (1 + rng.uniform(-0.3, 0.3)), 2),
            "1y": round(price * (1 + rng.uniform(-0.4, 0.5)), 2),
            "5d": round(price * (1 + rng.uniform(-0.03, 0.03)), 2),
            "ytd": round(price * (1 + rng.uniform(-0.25, 0.35)), 2),
            "change_pct": round(rng.uniform(-4, 4), 2),
            "change_dollar": round(price * rng.uniform(-0.04, 0.04), 2),
            "pe": round(rng.uniform(8, 65), 1),
            "debt_equity": round(rng.uniform(0.1, 2.5), 2),
            "ebit_margin": round(rng.uniform(5, 45), 1),
            "industry": rng.choice(industries),
            "earnings_date": "2026-11-15",
        }
    return data


def pct_change(current, prior):
    if current is None or prior is None or prior == 0:
        return None
    return (current - prior) / prior * 100


def build_rows(data):
    rows = []
    for sym in get_tickers():
        d = data.get(sym, {})
        price = d.get("price")
        rows.append({
            "symbol": sym, "price": price,
            "change_dollar": d.get("change_dollar"), "change_pct": d.get("change_pct"),
            "pe": d.get("pe"), "debt_equity": d.get("debt_equity"),
            "ebit_margin": d.get("ebit_margin"), "industry": d.get("industry"),
            "earnings_date": d.get("earnings_date"),
            "5d": pct_change(price, d.get("5d")) if price is not None else None,
            "1w": pct_change(price, d.get("1w")) if price is not None else None,
            "1m": pct_change(price, d.get("1m")) if price is not None else None,
            "6m": pct_change(price, d.get("6m")) if price is not None else None,
            "ytd": pct_change(price, d.get("ytd")) if price is not None else None,
            "1y": pct_change(price, d.get("1y")) if price is not None else None,
        })
    return rows


HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>KWP Terminal &mdash; Porter</title>
<style>
  :root {
    --bg: #000000; --panel: #0a0a0a; --line: #262626; --amber: #ffb000;
    --white: #d8d8d8; --green: #00d964; --red: #ff3b3b;
    --mono: "Menlo", "Consolas", "SFMono-Regular", monospace;
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--bg); color: var(--white); font-family: var(--mono); }
  header { display: flex; justify-content: space-between; align-items: center; padding: 10px 16px; border-bottom: 1px solid var(--line); flex-wrap: wrap; gap: 8px; }
  .brand { color: var(--amber); font-size: 14px; letter-spacing: 2px; }
  nav.fkeys { display: flex; gap: 2px; }
  nav.fkeys a { color: var(--white); text-decoration: none; font-size: 11px; padding: 5px 10px; border: 1px solid var(--line); background: var(--panel); }
  nav.fkeys a.active { color: #000; background: var(--amber); border-color: var(--amber); }
  .as-of { font-size: 10px; color: #666; }
  .wrap { padding: 14px 16px 40px; max-width: 1200px; margin: 0 auto; }
  .panel { background: var(--panel); border: 1px solid var(--line); }
  .panel h3 { margin: 0; padding: 6px 10px; font-size: 10.5px; color: #000; background: var(--amber); text-transform: uppercase; letter-spacing: 1px; font-weight: normal; }
  .p-scroll { overflow-x: auto; }
  .p-row {
    display: grid; grid-template-columns: 65px 80px 80px 65px 60px 65px 65px 65px 65px 65px 65px 65px 65px 240px 90px;
    gap: 8px; padding: 7px 12px; font-size: 12px; border-top: 1px solid var(--line);
    align-items: center; min-width: 1550px;
  }
  .p-row:first-child { border-top: none; }
  .p-row.p-head { color: #888; font-size: 9.5px; text-transform: uppercase; letter-spacing: 0.5px; background: #111; }
  .sortable { cursor: pointer; user-select: none; }
  .sortable:hover { color: var(--amber); }
  .p-symbol { color: var(--amber); }
  .p-num { text-align: right; }
  .p-industry { font-size: 11px; color: #bbb; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .t-up { color: var(--green); }
  .t-down { color: var(--red); }
  .t-na { color: #555; font-size: 11px; }
  .note { font-size: 11px; color: #666; margin-top: 12px; }
  .note b { color: #999; }
  footer { text-align: center; font-size: 10px; color: #444; margin-top: 24px; }
</style>
</head>
<body>
  <header>
    <div class="brand">KWP TERMINAL</div>
    <nav class="fkeys">
      <a href="index.html">[F1] MACRO</a>
      <a href="portfolio.html">[F2] PORTFOLIO</a>
      <a href="terminal.html">[F3] TERMINAL</a>
      <a href="porter.html" class="active">[F4] PORTER</a>
    </nav>
    <div class="as-of">__GENDATE__</div>
  </header>

  <div class="wrap">
    <div class="panel">
      <h3>Porter View</h3>
      <div class="p-scroll"><div id="holdings"></div></div>
    </div>
    <div class="note">P/E, Change $, and the Perf columns are live from GOOGLEFINANCE. D/E, EBIT Margin, Industry, and Earnings Date are your own manually-entered numbers (columns Q-T) — none of those are available from GOOGLEFINANCE at all.</div>
    __MISSING_WARNING__
    <footer>Prices via GOOGLEFINANCE &middot; generated __GENDATE__</footer>
  </div>

<script>
const ROWS = __ROWS__;

const SORT_COLUMNS = [
  { key: "symbol", label: "Symbol", type: "text" },
  { key: "price", label: "Price", type: "num" },
  { key: "change_dollar", label: "Change", type: "num" },
  { key: "change_pct", label: "Change %", type: "num" },
  { key: "pe", label: "P/E", type: "num" },
  { key: "5d", label: "5D", type: "num" },
  { key: "1w", label: "1W", type: "num" },
  { key: "1m", label: "1M", type: "num" },
  { key: "6m", label: "6M", type: "num" },
  { key: "ytd", label: "YTD", type: "num" },
  { key: "1y", label: "1Y", type: "num" },
  { key: "debt_equity", label: "D/E", type: "num" },
  { key: "ebit_margin", label: "EBIT %", type: "num" },
  { key: "industry", label: "Industry", type: "text" },
  { key: "earnings_date", label: "Earnings", type: "text" },
];
let sortKey = "symbol";
let sortDir = "asc";

function pctHtml(pct) {
  if (pct === null || pct === undefined) return '<span class="t-na">N/A</span>';
  const cls = pct >= 0 ? "t-up" : "t-down";
  const sign = pct >= 0 ? "+" : "";
  return `<span class="${cls}">${sign}${pct.toFixed(1)}%</span>`;
}
function dollarHtml(v) {
  if (v === null || v === undefined) return '<span class="t-na">N/A</span>';
  const cls = v >= 0 ? "t-up" : "t-down";
  const sign = v >= 0 ? "+" : "";
  return `<span class="${cls}">${sign}${v.toFixed(2)}</span>`;
}

function render() {
  const el = document.getElementById("holdings");
  const head = `<div class="p-row p-head">` + SORT_COLUMNS.map(col => {
    const active = col.key === sortKey;
    const arrow = active ? (sortDir === "desc" ? " \u25BC" : " \u25B2") : "";
    const cls = col.type === "text" ? "sortable" : "p-num sortable";
    return `<div class="${cls}" data-sort="${col.key}">${col.label}${arrow}</div>`;
  }).join("") + `</div>`;

  const sorted = [...ROWS].sort((a, b) => {
    if (a.price === null && b.price === null) return 0;
    if (a.price === null) return 1;
    if (b.price === null) return -1;
    const col = SORT_COLUMNS.find(c => c.key === sortKey);
    const va = a[col.key], vb = b[col.key];
    if (va === null || va === undefined) return 1;
    if (vb === null || vb === undefined) return -1;
    if (col.type === "text") return sortDir === "asc" ? va.localeCompare(vb) : vb.localeCompare(va);
    return sortDir === "asc" ? va - vb : vb - va;
  });

  const rows = sorted.map(r => {
    if (r.price === null) {
      return `<div class="p-row"><div class="p-symbol">${r.symbol}</div><div class="p-num t-na">no quote</div>${'<div class="p-num">&mdash;</div>'.repeat(13)}</div>`;
    }
    return `<div class="p-row">
      <div class="p-symbol">${r.symbol}</div>
      <div class="p-num">$${r.price.toFixed(2)}</div>
      <div class="p-num">${dollarHtml(r.change_dollar)}</div>
      <div class="p-num">${pctHtml(r.change_pct)}</div>
      <div class="p-num">${r.pe !== null && r.pe !== undefined ? r.pe.toFixed(1) : '<span class="t-na">N/A</span>'}</div>
      <div class="p-num">${pctHtml(r["5d"])}</div>
      <div class="p-num">${pctHtml(r["1w"])}</div>
      <div class="p-num">${pctHtml(r["1m"])}</div>
      <div class="p-num">${pctHtml(r["6m"])}</div>
      <div class="p-num">${pctHtml(r.ytd)}</div>
      <div class="p-num">${pctHtml(r["1y"])}</div>
      <div class="p-num">${r.debt_equity !== null && r.debt_equity !== undefined ? r.debt_equity.toFixed(2) : '<span class="t-na">N/A</span>'}</div>
      <div class="p-num">${r.ebit_margin !== null && r.ebit_margin !== undefined ? r.ebit_margin.toFixed(1) + '%' : '<span class="t-na">N/A</span>'}</div>
      <div class="p-industry">${r.industry || '<span class="t-na">N/A</span>'}</div>
      <div class="p-num">${r.earnings_date || '<span class="t-na">N/A</span>'}</div></div>`;
  }).join("");

  el.innerHTML = head + rows;
}

render();

document.getElementById("holdings").addEventListener("click", e => {
  const cell = e.target.closest(".sortable");
  if (!cell) return;
  const key = cell.dataset.sort;
  if (key === sortKey) {
    sortDir = sortDir === "desc" ? "asc" : "desc";
  } else {
    sortKey = key;
    sortDir = key === "symbol" ? "asc" : "desc";
  }
  render();
});
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser(description="Build porter.html.")
    ap.add_argument("-o", "--output", default="porter.html")
    ap.add_argument("--sheet-csv-url", default=SHEET_CSV_URL)
    ap.add_argument("--xlsx", default=None)
    ap.add_argument("--demo", action="store_true")
    args = ap.parse_args()

    if args.demo:
        data = demo_data()
    elif args.xlsx:
        data = read_xlsx_data(args.xlsx)
        pv.load_holdings_xlsx(args.xlsx)  # keeps the ticker list in sync with the Holdings tab
    else:
        data = fetch_csv_data(args.sheet_csv_url)
        pv.load_holdings_csv(pv.HOLDINGS_CSV_URL)

    rows = build_rows(data)
    missing = [r["symbol"] for r in rows if r["price"] is None]

    html = HTML_TEMPLATE.replace("__ROWS__", json.dumps(rows))
    html = html.replace("__GENDATE__", datetime.today().strftime("%Y-%m-%d %H:%M"))
    warning = ""
    if missing:
        warning = f'<div class="note" style="color:var(--red)">No quote for: {", ".join(missing)} — columns M-P may not be set up yet, or the ticker needs an exchange prefix.</div>'
    html = html.replace("__MISSING_WARNING__", warning)

    with open(args.output, "w") as f:
        f.write(html)
    print(f"Wrote {args.output} ({len(rows)} rows, {len(missing)} missing quotes)")


if __name__ == "__main__":
    main()
