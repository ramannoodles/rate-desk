"""
Terminal — visualizer
=======================
A dense, single-screen Bloomberg-style view of your macro data, in
monospace number-grid panels, black background, minimal chrome, and a
single period selector (1W / 1M / 6M / YTD / 1Y) that drives every panel
at once.

STOCK INFO IS FULLY BACK when portfolio data is provided — ticker,
Portfolio Value strip, and Top/Bottom Movers all show, and all of them
— including the ticker — follow the same period bar (Daily / 1W / 1M /
6M / YTD / 1Y) at the top of the page. Macro grids ignore "Daily" (FRED
has no daily reading) and just show a blank delta for that selection,
same honest-gap pattern as everywhere else in this script.

THE TICKER NEEDS THE changepct SHEET COLUMN to show real numbers on
"Daily" — see portfolio_visualize.py's docstring, column H. Without it,
Daily shows "N/A" for every symbol rather than a stale or wrong number;
every other period (1W/1M/6M/YTD) already works without that column.

This imports functions directly from macro_visualize.py and
portfolio_visualize.py rather than re-reading files a second way, so it
stays in sync with any changes to either script.

A NOTE ON "1W" FOR MACRO DATA: FRED macro series here are monthly, so
there's no real weekly reading to show — 1W and 1M will display the same
number. That's an honest reflection of the data's actual granularity,
not a bug.

USAGE
-----
python terminal_visualize.py macro_dashboard.xlsx -o terminal.html                    # macro only (current default)
python terminal_visualize.py macro_dashboard.xlsx --xlsx KWP.xlsx -o terminal.html    # macro + portfolio
python terminal_visualize.py macro_dashboard.xlsx --sheet-csv-url "https://..." -o terminal.html
python terminal_visualize.py macro_dashboard.xlsx --demo -o terminal.html
"""

import argparse
import json
import os
from datetime import datetime

import macro_visualize as mv
import portfolio_visualize as pv

# Indicators pulled out for the number-grid panels, by (sheet, display_name)
# — must match names in the workbook exactly (same names macro_visualize.py uses).
# Covers all 36 series in the workbook now, across 7 panels.
RATES_PANEL = [
    ("Leading Indicators", "Fed Funds Rate"),
    ("Leading Indicators", "Prime Rate"),
    ("Leading Indicators", "3M Treasury Yield"),
    ("Leading Indicators", "2Y Treasury Yield"),
    ("Leading Indicators", "5Y Treasury Yield"),
    ("Leading Indicators", "10Y Treasury Yield"),
    ("Leading Indicators", "30Y Treasury Yield"),
    ("Leading Indicators", "10Y-3M Yield Spread"),
    ("Leading Indicators", "10Y-2Y Yield Spread"),
]
INFLATION_PANEL = [
    ("Lagging Indicators", "CPI YoY"),
    ("Lagging Indicators", "Core CPI YoY"),
    ("Lagging Indicators", "Supercore CPI Proxy"),
    ("Lagging Indicators", "Core PCE Deflator"),
    ("Lagging Indicators", "PPI Final Demand"),
    ("Lagging Indicators", "5Y Breakeven Inflation"),
    ("Lagging Indicators", "10Y Breakeven Inflation"),
]
LABOR_PANEL = [
    ("Coincident Indicators", "Unemployment Rate"),
    ("Coincident Indicators", "Nonfarm Payrolls"),
    ("Coincident Indicators", "Avg Hourly Earnings"),
    ("Coincident Indicators", "Avg Weekly Hours"),
    ("Coincident Indicators", "JOLTS Job Openings"),
    ("Coincident Indicators", "Personal Income"),
]
RISK_PANEL = [
    ("Coincident Indicators", "Sahm Rule Recession Ind."),
    ("Leading Indicators", "Chicago Fed Fin. Conditions"),
]
MONEY_HOUSING_PANEL = [
    ("Leading Indicators", "M2 Money Supply"),
    ("Leading Indicators", "Housing Starts"),
    ("Leading Indicators", "Building Permits"),
]
BUSINESS_PANEL = [
    ("Business Activity", "Mfg New Orders"),
    ("Business Activity", "Durable Goods Orders"),
    ("Business Activity", "Industrial Production"),
    ("Business Activity", "Industrial Prod: Mfg"),
    ("Business Activity", "Retail Sales"),
]
GROWTH_PANEL = [
    ("Lagging Indicators", "Real GDP YoY"),
    ("Lagging Indicators", "WTI Crude Oil"),
    ("Coincident Indicators", "UMich Consumer Sentiment"),
    ("Coincident Indicators", "Personal Spending"),
]

PERIODS = ["1w", "1m", "6m", "ytd", "1y"]
PERIOD_LABELS = {"1w": "1W", "1m": "1M", "6m": "6M", "ytd": "YTD", "1y": "1Y"}


def _macro_periods(dates, values):
    """Point-change AND %-change vs each lookback, for a monthly series.
    Point change reads better than % alone for rate-type series (Fed
    Funds, CPI YoY are already percentages), so both are shown together —
    e.g. "3.63  \u25B2 0.11 (+3.1%)"."""
    latest = latest_idx = None
    for i in range(len(values) - 1, -1, -1):
        if values[i] is not None:
            latest, latest_idx = values[i], i
            break
    out = {"latest": latest}
    if latest is None:
        return {**out, **{p: {"pt": None, "pct": None} for p in PERIODS}}

    def back(k):
        idx = latest_idx - k
        if idx < 0 or values[idx] is None:
            return {"pt": None, "pct": None}
        prior = values[idx]
        pt = latest - prior
        pct = None if prior == 0 else pt / abs(prior) * 100
        return {"pt": pt, "pct": pct}

    out["1w"] = back(1)   # monthly data: same as 1m, see module docstring
    out["1m"] = back(1)
    out["6m"] = back(6)
    out["1y"] = back(12)
    # YTD: vs the first data point in the latest year present.
    latest_year = dates[latest_idx].split(" ")[-1]
    jan_idx = next((i for i, d in enumerate(dates) if d.startswith("Jan") and d.endswith(latest_year)), None)
    if jan_idx is None or values[jan_idx] is None:
        out["ytd"] = {"pt": None, "pct": None}
    else:
        prior = values[jan_idx]
        pt = latest - prior
        out["ytd"] = {"pt": pt, "pct": None if prior == 0 else pt / abs(prior) * 100}
    return out


def _grid_data(series_map, refs):
    rows = []
    for sheet, name in refs:
        s = series_map.get((sheet, name))
        periods = _macro_periods(s["dates"], s["values"]) if s else {"latest": None, **{p: None for p in PERIODS}}
        rows.append({"name": name, **periods})
    return rows


def build_terminal_html(macro_path, portfolio_data=None, insights_path="insights.json"):
    summary, series = mv.read_workbook(macro_path)
    grids = {
        "rates": _grid_data(series, RATES_PANEL),
        "inflation": _grid_data(series, INFLATION_PANEL),
        "labor": _grid_data(series, LABOR_PANEL),
        "risk": _grid_data(series, RISK_PANEL),
        "money_housing": _grid_data(series, MONEY_HOUSING_PANEL),
        "business": _grid_data(series, BUSINESS_PANEL),
        "growth": _grid_data(series, GROWTH_PANEL),
    }

    has_portfolio = portfolio_data is not None
    total_value = None
    positions = []
    if has_portfolio:
        port_rows = pv.build_rows(portfolio_data)
        priced = [r for r in port_rows if r["value"] is not None]
        total_value = sum(r["value"] for r in priced) + pv.CASH_VALUE
        positions = [
            {
                "symbol": r["symbol"], "price": r["price"],
                "1d": r["1d"], "1w": r["1w"], "1m": r["1m"], "6m": r["6m"],
                "1y": r["1y"],
                # portfolio_visualize.py now computes real YTD from the
                # ytd_close column (column P) — was hardcoded None here
                # from before that column existed. Real data now.
                "ytd": r["ytd"],
            }
            for r in priced
        ]

    html = HTML_TEMPLATE
    html = html.replace("__TICKER_BLOCK__", _ticker_block_html() if has_portfolio else "")
    html = html.replace("__TOTAL_BLOCK__", _total_block_html(total_value) if has_portfolio else "")
    html = html.replace("__MOVERS_BLOCK__", _movers_block_html() if has_portfolio else "")
    html = html.replace("__GRIDS__", json.dumps(grids))
    html = html.replace("__POSITIONS__", json.dumps(positions))
    html = html.replace("__HAS_PORTFOLIO__", json.dumps(has_portfolio))
    html = html.replace("__CALENDAR__", mv.build_calendar_html())
    html = html.replace("__INSIGHTS_JSON__", json.dumps(build_insights_json(insights_path, include_portfolio=has_portfolio)))
    html = html.replace("__GENDATE__", datetime.today().strftime("%Y-%m-%d %H:%M"))
    return html


def _ticker_block_html():
    return '<div class="ticker-wrap"><div class="ticker" id="ticker"></div></div>'


def _total_block_html(total_value):
    return (
        '<div class="total-strip"><div><div class="total-label">Portfolio Value</div>'
        f'<div class="total-value">${total_value:,.0f}</div></div></div>'
    )


def _movers_block_html():
    return (
        '<div class="panel"><h3 id="winners-h3">Top Movers</h3><div id="winners"></div></div>\n'
        '<div class="panel"><h3 id="losers-h3">Bottom Movers</h3><div id="losers"></div></div>'
    )


def build_insights_json(insights_path, include_portfolio=True):
    if not os.path.exists(insights_path):
        return []
    with open(insights_path) as f:
        items = json.load(f)
    # Backward-compat: older insights.json files (before the Daily toggle
    # existed) have no "period"/"source" tag.
    for it in items:
        it.setdefault("period", "weekly")
        it.setdefault("key", "")
        it.setdefault("source", "portfolio" if it["key"].startswith("portfolio") else "macro")
    if not include_portfolio:
        items = [it for it in items if it["source"] == "macro"]
    return items


HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>KWP Terminal</title>
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
  .ticker-wrap { background: #111; border-bottom: 1px solid var(--line); overflow: hidden; white-space: nowrap; padding: 6px 0; }
  .ticker { display: inline-block; animation: scroll 60s linear infinite; }
  .tick { display: inline-block; padding: 0 22px; font-size: 12px; color: var(--white); }
  .tick b { color: var(--amber); font-weight: normal; }
  @keyframes scroll { from { transform: translateX(0); } to { transform: translateX(-33.333%); } }

  header { display: flex; justify-content: space-between; align-items: center; padding: 10px 16px; border-bottom: 1px solid var(--line); flex-wrap: wrap; gap: 8px; }
  .brand { color: var(--amber); font-size: 14px; letter-spacing: 2px; }
  nav.fkeys { display: flex; gap: 2px; }
  nav.fkeys a {
    color: var(--white); text-decoration: none; font-size: 11px; padding: 5px 10px;
    border: 1px solid var(--line); background: var(--panel);
  }
  nav.fkeys a.active { color: #000; background: var(--amber); border-color: var(--amber); }
  .as-of { font-size: 10px; color: #666; }

  .period-bar { display: flex; gap: 2px; padding: 10px 16px; border-bottom: 1px solid var(--line); align-items: center; }
  .period-label { font-size: 10px; color: #888; text-transform: uppercase; letter-spacing: 1px; margin-right: 10px; }
  .period-bar button {
    font-family: var(--mono); font-size: 11px; color: var(--white); background: var(--panel);
    border: 1px solid var(--line); padding: 5px 14px; cursor: pointer;
  }
  .period-bar button.active { color: #000; background: var(--amber); border-color: var(--amber); }

  .wrap { padding: 14px 16px 40px; max-width: 1400px; margin: 0 auto; }
  .total-strip { display: flex; gap: 30px; align-items: baseline; margin-bottom: 16px; padding-bottom: 12px; border-bottom: 1px solid var(--line); }
  .total-label { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 1px; }
  .total-value { font-size: 26px; color: var(--amber); }

  .grid-wrap { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; margin-bottom: 10px; }
  .panel { background: var(--panel); border: 1px solid var(--line); }
  .panel h3 {
    margin: 0; padding: 6px 10px; font-size: 10.5px; color: #000; background: var(--amber);
    text-transform: uppercase; letter-spacing: 1px; font-weight: normal;
  }
  .t-row { display: flex; justify-content: space-between; padding: 5px 10px; font-size: 12px; border-top: 1px solid var(--line); }
  .t-row:first-child { border-top: none; }
  .t-name { color: #aaa; }
  .t-val { color: var(--white); text-align: right; }
  .t-delta { min-width: 110px; text-align: right; font-size: 11px; }
  .t-up { color: var(--green); }
  .t-down { color: var(--red); }
  .t-na { color: #555; font-size: 10px; padding: 10px; text-align: center; }

  .cal-strip { display: flex; gap: 8px; overflow-x: auto; margin-top: 4px; }
  .cal-card { border: 1px solid var(--line); background: var(--panel); padding: 6px 10px; font-size: 11px; white-space: nowrap; flex: 0 0 auto; }
  .cal-card.fomc { border-color: var(--amber); }
  .cal-card .cal-date { color: var(--amber); }
  .cal-card .cal-name { color: var(--white); }
  .cal-card .cal-time { color: #888; }
  .cal-empty { font-size: 11px; color: #666; }

  .insight-card { border-top: 1px solid var(--line); padding: 10px 12px; font-size: 12px; }
  .insight-card:first-child { border-top: none; }
  .insight-head { margin-bottom: 4px; }
  .insight-label { color: var(--amber); }
  .insight-when { color: #666; font-size: 10px; float: right; }
  .insight-text { color: #ccc; line-height: 1.5; }

  .insights-head { display: flex; align-items: center; background: var(--amber); padding: 0 10px; }
  .insights-head h3 { background: transparent; padding: 6px 0; }
  .insights-head .btn-group { border-color: #000; }
  .insights-head .btn-group button {
    font-family: var(--mono); font-size: 10px; color: #000; background: transparent;
    border: 1px solid #000; padding: 3px 10px; cursor: pointer;
  }
  .insights-head .btn-group button.active { background: #000; color: var(--amber); }

  footer { text-align: center; font-size: 10px; color: #444; margin-top: 24px; }
  @media (max-width: 900px) { .grid-wrap { grid-template-columns: 1fr; } }
</style>
</head>
<body>
  __TICKER_BLOCK__

  <header>
    <div class="brand">KWP TERMINAL</div>
    <nav class="fkeys">
      <a href="index.html">[F1] MACRO</a>
      <a href="portfolio.html">[F2] PORTFOLIO</a>
      <a href="terminal.html" class="active">[F3] TERMINAL</a>
      <a href="porter.html">[F4] PORTER</a>
    </nav>
    <div class="as-of">__GENDATE__</div>
  </header>

  <div class="period-bar" id="period-bar">
    <span class="period-label">Period</span>
    <button data-period="1d">Daily</button>
    <button data-period="1w" class="active">1W</button>
    <button data-period="1m">1M</button>
    <button data-period="6m">6M</button>
    <button data-period="ytd">YTD</button>
    <button data-period="1y">1Y</button>
  </div>

  <div class="wrap">
    __TOTAL_BLOCK__

    <div class="grid-wrap">
      <div class="panel"><h3>Rates</h3><div id="grid-rates"></div></div>
      <div class="panel"><h3>Inflation</h3><div id="grid-inflation"></div></div>
      <div class="panel"><h3>Labor</h3><div id="grid-labor"></div></div>
      <div class="panel"><h3>Risk Signals</h3><div id="grid-risk"></div></div>
      <div class="panel"><h3>Money &amp; Housing</h3><div id="grid-money_housing"></div></div>
      <div class="panel"><h3>Business Activity</h3><div id="grid-business"></div></div>
      <div class="panel"><h3>Growth &amp; Consumer</h3><div id="grid-growth"></div></div>
      __MOVERS_BLOCK__
    </div>

    <div class="panel">
      <h3>Next 3 Weeks</h3>
      <div class="cal-strip">__CALENDAR__</div>
    </div>

    <div class="panel" style="margin-top: 10px;">
      <div class="insights-head">
        <h3 style="flex:1">Insights &mdash; Drastic Moves</h3>
        <div class="btn-group" id="insight-period-group">
          <button data-iperiod="daily">Daily</button>
          <button data-iperiod="weekly" class="active">Past Week</button>
        </div>
      </div>
      <div id="insights"></div>
    </div>

    <footer>DATA: FRED &middot; GOOGLE FINANCE &middot; NOT INVESTMENT ADVICE</footer>
  </div>

<script>
const GRIDS = __GRIDS__;
const POSITIONS = __POSITIONS__;
const HAS_PORTFOLIO = __HAS_PORTFOLIO__;
const INSIGHTS = __INSIGHTS_JSON__;
const PERIOD_LABEL = {"1d": "Daily", "1w": "1W", "1m": "1M", "6m": "6M", "ytd": "YTD", "1y": "1Y"};
let currentPeriod = "1w";
let currentInsightPeriod = "weekly";

function fmtDelta(d) {
  if (!d || d.pt === null || d.pt === undefined) return "";
  const cls = d.pt >= 0 ? "t-up" : "t-down";
  const arrow = d.pt >= 0 ? "\u25B2" : "\u25BC";
  const pctTxt = (d.pct === null || d.pct === undefined) ? "" : ` (${d.pct >= 0 ? "+" : ""}${d.pct.toFixed(1)}%)`;
  return `<span class="${cls}">${arrow} ${Math.abs(d.pt).toFixed(2)}${pctTxt}</span>`;
}

function renderGrid(elId, rows) {
  const el = document.getElementById(elId);
  el.innerHTML = rows.map(r => {
    if (r.latest === null) {
      return `<div class="t-row"><div class="t-name">${r.name}</div><div class="t-val">N/A</div><div class="t-delta"></div></div>`;
    }
    return `<div class="t-row"><div class="t-name">${r.name}</div>
      <div class="t-val">${r.latest.toFixed(2)}</div>
      <div class="t-delta">${fmtDelta(r[currentPeriod])}</div></div>`;
  }).join("");
}

function renderMovers() {
  if (!HAS_PORTFOLIO) return;
  document.getElementById("winners-h3").textContent = `Top Movers (${PERIOD_LABEL[currentPeriod]} %)`;
  document.getElementById("losers-h3").textContent = `Bottom Movers (${PERIOD_LABEL[currentPeriod]} %)`;

  const ranked = POSITIONS.filter(p => p[currentPeriod] !== null && p[currentPeriod] !== undefined)
                           .sort((a, b) => b[currentPeriod] - a[currentPeriod]);

  if (ranked.length === 0) {
    const msg = `<div class="t-na">No ${PERIOD_LABEL[currentPeriod]} data tracked yet for positions</div>`;
    document.getElementById("winners").innerHTML = msg;
    document.getElementById("losers").innerHTML = msg;
    return;
  }

  const rowHtml = p => {
    const pct = p[currentPeriod];
    const cls = pct >= 0 ? "t-up" : "t-down";
    const sign = pct >= 0 ? "+" : "";
    return `<div class="t-row"><div class="t-name">${p.symbol}</div>
      <div class="t-val">$${p.price.toFixed(2)}</div>
      <div class="t-delta"><span class="${cls}">${sign}${pct.toFixed(1)}%</span></div></div>`;
  };
  document.getElementById("winners").innerHTML = ranked.slice(0, 5).map(rowHtml).join("");
  document.getElementById("losers").innerHTML = ranked.slice(-5).reverse().map(rowHtml).join("");
}

function renderTicker() {
  if (!HAS_PORTFOLIO) return;
  const items = POSITIONS.map(p => {
    const pct = p[currentPeriod];
    const cls = pct !== null && pct !== undefined && pct >= 0 ? "t-up" : "t-down";
    const sign = pct !== null && pct !== undefined && pct >= 0 ? "+" : "";
    const txt = (pct === null || pct === undefined) ? "N/A" : `${sign}${pct.toFixed(1)}%`;
    return `<span class="tick">${p.symbol} <b>${p.price.toFixed(2)}</b> <span class="${cls}">${txt}</span></span>`;
  }).join("");
  document.getElementById("ticker").innerHTML = items.repeat(3);
}

function renderInsights() {
  const items = INSIGHTS.filter(it => it.period === currentInsightPeriod);
  const el = document.getElementById("insights");

  if (items.length === 0) {
    const why = currentInsightPeriod === "daily"
      ? "No daily insights yet \u2014 your Google Sheet needs a changepct column (see portfolio_visualize.py's docstring, column H) and generate_insights.py needs to be run again after that."
      : "Nothing crossed the drastic-move threshold on the last check.";
    el.innerHTML = `<div class="cal-empty">${why}</div>`;
    return;
  }

  el.innerHTML = items.map(it => {
    const cls = it.pct >= 0 ? "t-up" : "t-down";
    const sign = it.pct >= 0 ? "+" : "";
    const move = (it.pt !== null && it.pt !== undefined)
      ? `${it.pt >= 0 ? "+" : ""}${it.pt.toFixed(2)} (${sign}${it.pct.toFixed(1)}%)`
      : `${sign}${it.pct.toFixed(1)}%`;
    return `<div class="insight-card"><div class="insight-head">
      <span class="insight-label">${it.label}</span>
      <span class="${cls}">${move}</span>
      <span class="insight-when">${it.period_label} &middot; ${it.generated_at}</span></div>
      <div class="insight-text">${it.text}</div></div>`;
  }).join("");
}

function renderAll() {
  renderGrid("grid-rates", GRIDS.rates);
  renderGrid("grid-inflation", GRIDS.inflation);
  renderGrid("grid-labor", GRIDS.labor);
  renderGrid("grid-risk", GRIDS.risk);
  renderGrid("grid-money_housing", GRIDS.money_housing);
  renderGrid("grid-business", GRIDS.business);
  renderGrid("grid-growth", GRIDS.growth);
  renderMovers();
  renderTicker();
  renderInsights();
}

document.getElementById("period-bar").addEventListener("click", e => {
  const btn = e.target.closest("button[data-period]");
  if (!btn) return;
  currentPeriod = btn.dataset.period;
  document.querySelectorAll("#period-bar button").forEach(b => b.classList.toggle("active", b === btn));
  renderAll();
});

document.getElementById("insight-period-group").addEventListener("click", e => {
  const btn = e.target.closest("button[data-iperiod]");
  if (!btn) return;
  currentInsightPeriod = btn.dataset.iperiod;
  document.querySelectorAll("#insight-period-group button").forEach(b => b.classList.toggle("active", b === btn));
  renderInsights();
});

renderAll();
</script>
</body>
</html>
"""


def main():
    ap = argparse.ArgumentParser(description="Build the Bloomberg-style terminal.html page.")
    ap.add_argument("macro_workbook")
    ap.add_argument("-o", "--output", default="terminal.html")
    ap.add_argument("--sheet-csv-url", default=None)
    ap.add_argument("--xlsx", default=None)
    ap.add_argument("--demo", action="store_true")
    args = ap.parse_args()

    if args.demo:
        portfolio_data = pv.demo_price_data()
    elif args.xlsx:
        portfolio_data = pv.read_xlsx_price_data(args.xlsx)
        pv.load_holdings_xlsx(args.xlsx)  # picks up shares/cost/cash if a Holdings tab exists
    elif args.sheet_csv_url:
        portfolio_data = pv.fetch_price_data(args.sheet_csv_url)
        pv.load_holdings_csv(pv.HOLDINGS_CSV_URL)
    elif pv.SHEET_CSV_URL and "PASTE_YOUR" not in pv.SHEET_CSV_URL:
        # No flag passed, but portfolio_visualize.py's own URL has been
        # set for real — reuse it automatically so there's only ever one
        # place to update the URL, not one per script.
        portfolio_data = pv.fetch_price_data(pv.SHEET_CSV_URL)
        pv.load_holdings_csv(pv.HOLDINGS_CSV_URL)
    else:
        # No portfolio source anywhere -> macro-only page. No error: this
        # is the current default while stock info is off for now (see
        # module docstring). Pass --xlsx/--sheet-csv-url/--demo to bring
        # the ticker, movers, and portfolio value strip back.
        portfolio_data = None

    html = build_terminal_html(args.macro_workbook, portfolio_data)
    with open(args.output, "w") as f:
        f.write(html)
    mode = "with portfolio data" if portfolio_data is not None else "macro-only"
    print(f"Wrote {args.output} ({mode})")


if __name__ == "__main__":
    main()
