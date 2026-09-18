"""
Macro Dashboard — visualizer
=============================
Reads a workbook shaped like macro_dashboard.xlsx (a "Dashboard" summary
tab plus "Leading Indicators" / "Business Activity" / "Coincident
Indicators" / "Lagging Indicators" tabs, each with FRED ID / Display Name
/ Unit / Signal Logic in columns A-D and monthly values from column E
onward) and produces a single self-contained HTML dashboard:

  - a status board (all indicators, grouped by category, colored by
    the workbook's own computed Status column)
  - the same 11 grouped charts already defined as native Excel chart
    tabs in the source workbook (Yield Curve, Inversion, Real Rate,
    Inflation, Labor, M2, Breakevens, Industrial, Consumer, Sahm Rule,
    Financial Conditions) — rebuilt from the underlying series so they
    render in a browser instead of requiring Excel.

USAGE
-----
python macro_visualize.py path/to/macro_dashboard.xlsx -o dashboard.html

Re-run it against any future export from the same pipeline (e.g. your
macro_data_pull.py output) to refresh the dashboard.
"""

import argparse
import json
from datetime import date, datetime, timedelta

from openpyxl import load_workbook

# --- Release calendar --------------------------------------------------
# FRED has no forward-looking calendar or consensus data, so unlike the
# rest of this script, this list is hand-maintained. Refresh it from the
# primary sources periodically:
#   BLS (CPI, PPI, jobs, JOLTS): https://www.bls.gov/schedule/
#   BEA (GDP, PCE):              https://www.bea.gov/news/schedule
#   Census (Retail Sales):       https://www.census.gov/economic-indicators/calendar-listview.html
#   Fed (FOMC):                  https://www.federalreserve.gov/monetarypolicy/fomccalendars.htm
# (date, time_et, release_name, period_covered)
RELEASE_CALENDAR = [
    (date(2026, 9, 11), "8:30 AM ET", "CPI", "Aug 2026"),
    (date(2026, 9, 16), "8:30 AM ET", "Retail Sales", "Aug 2026"),
    (date(2026, 9, 16), "2:00 PM ET", "FOMC Rate Decision", "Sep 15-16 mtg"),
    (date(2026, 9, 18), "9:15 AM ET", "Industrial Production", "Aug 2026"),
    (date(2026, 9, 25), "8:30 AM ET", "Durable Goods Orders", "Aug 2026"),
    (date(2026, 9, 25), "10:00 AM ET", "UMich Consumer Sentiment (final, est.)", "Sep 2026"),
    (date(2026, 9, 29), "10:00 AM ET", "JOLTS", "Aug 2026"),
    (date(2026, 9, 30), "8:30 AM ET", "GDP (3rd est.)", "Q2 2026"),
    (date(2026, 9, 30), "8:30 AM ET", "PCE Inflation", "Aug 2026"),
    (date(2026, 10, 2), "8:30 AM ET", "Nonfarm Payrolls", "Sep 2026"),
    (date(2026, 10, 2), "10:00 AM ET", "Mfg New Orders (full M3 report)", "Aug 2026"),
    (date(2026, 10, 7), "2:00 PM ET", "FOMC Minutes", "Sep 15-16 mtg"),
    (date(2026, 10, 14), "8:30 AM ET", "CPI", "Sep 2026"),
    (date(2026, 10, 15), "8:30 AM ET", "Retail Sales", "Sep 2026"),
    (date(2026, 10, 15), "8:30 AM ET", "PPI", "Sep 2026"),
    (date(2026, 10, 16), "9:15 AM ET", "Industrial Production", "Sep 2026"),
    (date(2026, 10, 20), "8:30 AM ET", "Housing Starts & Building Permits", "Sep 2026"),
    (date(2026, 10, 27), "8:30 AM ET", "Durable Goods Orders", "Sep 2026"),
    (date(2026, 10, 28), "2:00 PM ET", "FOMC Rate Decision", "Oct 27-28 mtg"),
    (date(2026, 10, 29), "8:30 AM ET", "GDP (advance)", "Q3 2026"),
    (date(2026, 10, 29), "8:30 AM ET", "PCE Inflation", "Sep 2026"),
    (date(2026, 10, 30), "10:00 AM ET", "UMich Consumer Sentiment (final, est.)", "Oct 2026"),
    (date(2026, 11, 3), "10:00 AM ET", "Mfg New Orders (full M3 report)", "Sep 2026"),
    (date(2026, 11, 6), "8:30 AM ET", "Nonfarm Payrolls", "Oct 2026"),
    (date(2026, 11, 17), "8:30 AM ET", "Retail Sales", "Oct 2026"),
    (date(2026, 11, 17), "9:15 AM ET", "Industrial Production", "Oct 2026"),
    (date(2026, 11, 18), "8:30 AM ET", "Housing Starts & Building Permits", "Oct 2026"),
    (date(2026, 11, 18), "2:00 PM ET", "FOMC Minutes", "Oct 27-28 mtg"),
    (date(2026, 11, 25), "8:30 AM ET", "Durable Goods Orders", "Oct 2026"),
    (date(2026, 11, 27), "10:00 AM ET", "UMich Consumer Sentiment (final, est.)", "Nov 2026"),
]


def build_calendar_html(today=None):
    today = today or date.today()
    window_end = today + timedelta(days=21)
    upcoming = sorted(r for r in RELEASE_CALENDAR if today <= r[0] <= window_end)
    if not upcoming:
        return ('<div class="cal-empty">No major releases in the next two weeks on file — '
                'RELEASE_CALENDAR in macro_visualize.py needs fresh dates added.</div>')
    cards = []
    for d, time_str, name, period in upcoming:
        cls = "cal-card fomc" if "FOMC" in name else "cal-card"
        day_label = f"{d.strftime('%a %b')} {d.day}"
        cards.append(
            f'<div class="{cls}"><div class="cal-date">{day_label}</div>'
            f'<div class="cal-name">{name}</div>'
            f'<div class="cal-time">{time_str} &middot; {period}</div></div>'
        )
    return "\n".join(cards)


CATEGORY_SHEETS = [
    "Leading Indicators",
    "Business Activity",
    "Coincident Indicators",
    "Lagging Indicators",
]

# Mirrors the workbook's own C1-C11 chart tabs: (title, [(sheet, display_name), ...])
PANELS = [
    ("Yield Curve", [
        ("Leading Indicators", "3M Treasury Yield"),
        ("Leading Indicators", "2Y Treasury Yield"),
        ("Leading Indicators", "5Y Treasury Yield"),
        ("Leading Indicators", "10Y Treasury Yield"),
        ("Leading Indicators", "30Y Treasury Yield"),
    ]),
    ("3M / 10Y / 30Y Treasury Yields", [
        ("Leading Indicators", "3M Treasury Yield"),
        ("Leading Indicators", "10Y Treasury Yield"),
        ("Leading Indicators", "30Y Treasury Yield"),
    ]),
    ("Yield Curve Inversion (0 = trigger)", [
        ("Leading Indicators", "10Y-3M Yield Spread"),
        ("Leading Indicators", "10Y-2Y Yield Spread"),
    ]),
    ("Real Rate: Fed Funds vs CPI YoY", [
        ("Leading Indicators", "Fed Funds Rate"),
        ("Lagging Indicators", "CPI YoY"),
    ]),
    ("Inflation Measures", [
        ("Lagging Indicators", "CPI YoY"),
        ("Lagging Indicators", "Core CPI YoY"),
        ("Lagging Indicators", "Supercore CPI Proxy"),
        ("Lagging Indicators", "Core PCE Deflator"),
        ("Lagging Indicators", "PPI Final Demand"),
    ]),
    ("Labor Market", [
        ("Coincident Indicators", "Unemployment Rate"),
        ("Coincident Indicators", "Avg Weekly Hours"),
    ]),
    ("Employment & Income", [
        ("Coincident Indicators", "Nonfarm Payrolls"),
        ("Coincident Indicators", "Avg Hourly Earnings"),
        ("Coincident Indicators", "Personal Income"),
    ]),
    ("M2 Money Supply", [
        ("Leading Indicators", "M2 Money Supply"),
    ]),
    ("Inflation Breakevens", [
        ("Lagging Indicators", "5Y Breakeven Inflation"),
        ("Lagging Indicators", "10Y Breakeven Inflation"),
    ]),
    ("Industrial Activity", [
        ("Business Activity", "Industrial Production"),
        ("Business Activity", "Industrial Prod: Mfg"),
        ("Business Activity", "Durable Goods Orders"),
        ("Business Activity", "Mfg New Orders"),
    ]),
    ("Housing", [
        ("Leading Indicators", "Housing Starts"),
        ("Leading Indicators", "Building Permits"),
    ]),
    ("Consumer Activity", [
        ("Coincident Indicators", "UMich Consumer Sentiment"),
        ("Coincident Indicators", "Personal Spending"),
        ("Business Activity", "Retail Sales"),
    ]),
    ("Sahm Rule Recession Indicator (0.50 = trigger)", [
        ("Coincident Indicators", "Sahm Rule Recession Ind."),
    ]),
    ("Financial Conditions", [
        ("Leading Indicators", "Chicago Fed Fin. Conditions"),
        ("Coincident Indicators", "JOLTS Job Openings"),
    ]),
    ("Prime Rate vs Fed Funds", [
        ("Leading Indicators", "Prime Rate"),
        ("Leading Indicators", "Fed Funds Rate"),
    ]),
    ("Real GDP Growth (YoY)", [
        ("Lagging Indicators", "Real GDP YoY"),
    ]),
    ("WTI Crude Oil", [
        ("Lagging Indicators", "WTI Crude Oil"),
    ]),
]


def _fmt_date(v):
    if isinstance(v, datetime):
        return v.strftime("%b %Y")
    return str(v)


def read_workbook(path):
    wb = load_workbook(path, data_only=True)

    # --- Dashboard summary ---
    ws = wb["Dashboard"]
    header_row = None
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True), start=1):
        if row[0] == "Category":
            header_row = i
            break
    summary = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        category = row[0]
        if not category or not row[1]:  # skip blanks and the trailing notes row
            continue
        summary.append({
            "category": category,
            "id": row[1],
            "name": row[2],
            "unit": row[3],
            "latest_month": _fmt_date(row[5]),
            "latest": row[6],
            "prior": row[7],
            "change": row[8],
            "status": row[9],
        })

    # --- Category tabs: build {(sheet, display_name): {"dates":[...], "values":[...]}} ---
    series = {}
    for sheet_name in CATEGORY_SHEETS:
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        header_idx = next(i for i, r in enumerate(rows) if r[0] == "FRED ID")
        dates_row = rows[header_idx]
        dates = [_fmt_date(d) for d in dates_row[4:] if d is not None]
        for r in rows[header_idx + 1:]:
            if not r[0]:
                continue
            display_name = r[1]
            values = list(r[4:4 + len(dates)])
            series[(sheet_name, display_name)] = {"dates": dates, "values": values}

    return summary, series


HTML_TEMPLATE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>Macro Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<style>
  :root {
    --ink: #0f1a17;
    --paper: #f4f1e8;
    --panel: #16211d;
    --line: rgba(244,241,232,0.12);
    --amber: #d9a441;
    --up: #6a9e7f;
    --down: #c96a4a;
    --flat: rgba(244,241,232,0.45);
    --mono: "IBM Plex Mono", "SFMono-Regular", Menlo, monospace;
    --serif: "Source Serif Pro", Georgia, "Times New Roman", serif;
    --sans: "Inter", "Helvetica Neue", Arial, sans-serif;
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--ink); color: var(--paper); font-family: var(--sans); padding: 28px 20px 60px; }
  .wrap { max-width: 1180px; margin: 0 auto; }
  header { border-bottom: 1px solid var(--line); padding-bottom: 18px; margin-bottom: 26px; }
  h1 { font-family: var(--serif); font-weight: 600; font-size: 32px; margin: 0; }
  .sub { font-size: 13px; color: rgba(244,241,232,0.55); margin-top: 6px; }
  h2 { font-family: var(--serif); font-weight: 600; font-size: 19px; margin: 40px 0 14px; }

  .board { border: 1px solid var(--line); }
  .cat-group { border-bottom: 1px solid var(--line); }
  .cat-group:last-child { border-bottom: none; }
  .cat-head {
    background: var(--panel); padding: 10px 14px; font-family: var(--mono);
    font-size: 11.5px; color: var(--amber); letter-spacing: 0.3px;
  }
  .row {
    display: grid; grid-template-columns: 1fr 90px 90px 130px;
    gap: 8px; padding: 8px 14px; font-size: 13px; border-top: 1px solid var(--line);
    align-items: center;
  }
  .row.head { color: rgba(244,241,232,0.45); font-size: 11px; font-family: var(--mono); border-top: none; }
  .row .name { }
  .row .num { font-family: var(--mono); text-align: right; }
  .pill {
    font-family: var(--mono); font-size: 11px; padding: 2px 8px; border-radius: 2px;
    text-align: center; width: fit-content; justify-self: end;
  }
  .pill.up { background: rgba(106,158,127,0.18); color: var(--up); }
  .pill.down { background: rgba(201,106,74,0.18); color: var(--down); }
  .pill.flat { background: rgba(244,241,232,0.08); color: var(--flat); }

  .calendar-label { font-family: var(--mono); font-size: 11px; color: rgba(244,241,232,0.4); margin-bottom: 10px; letter-spacing: 0.3px; }
  .calendar { display: flex; gap: 10px; overflow-x: auto; margin-bottom: 34px; padding-bottom: 2px; }
  .cal-card {
    flex: 0 0 auto; min-width: 148px; background: var(--panel); border: 1px solid var(--line);
    border-left: 2px solid var(--line); padding: 10px 14px;
  }
  .cal-card.fomc { border-left: 2px solid var(--amber); }
  .cal-date { font-family: var(--mono); font-size: 11px; color: var(--amber); }
  .cal-name { font-size: 13px; margin-top: 4px; }
  .cal-time { font-family: var(--mono); font-size: 10.5px; color: rgba(244,241,232,0.4); margin-top: 3px; }
  .cal-empty { font-size: 12px; color: rgba(244,241,232,0.4); margin-bottom: 34px; }

  nav.tabs { display: flex; gap: 4px; margin-bottom: 22px; }
  nav.tabs a {
    font-family: var(--mono); font-size: 12px; text-decoration: none; color: rgba(244,241,232,0.5);
    padding: 8px 16px; border: 1px solid var(--line); border-bottom: none;
  }
  nav.tabs a.active { color: var(--amber); border-color: var(--amber); background: var(--panel); }

  .controls { display: flex; gap: 22px; align-items: center; margin-bottom: 18px; flex-wrap: wrap; }
  .btn-group { display: flex; border: 1px solid var(--line); border-radius: 3px; overflow: hidden; }
  .btn-group button {
    background: transparent; color: var(--paper); border: none; border-right: 1px solid var(--line);
    padding: 7px 14px; font-family: var(--mono); font-size: 11.5px; cursor: pointer; letter-spacing: 0.3px;
  }
  .btn-group button:last-child { border-right: none; }
  .btn-group button.active { background: var(--amber); color: #1a1206; }
  .btn-group-label { font-family: var(--mono); font-size: 11px; color: rgba(244,241,232,0.4); margin-right: 8px; }

  .panels { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  .panel { background: var(--panel); border: 1px solid var(--line); padding: 18px; }
  .panel h3 { font-family: var(--serif); font-weight: 600; font-size: 15px; margin: 0 0 12px; }
  canvas { max-height: 230px; }
  footer { text-align: center; font-size: 11px; color: rgba(244,241,232,0.4); margin-top: 34px; font-family: var(--mono); }
  @media (max-width: 900px) {
    .panels { grid-template-columns: 1fr; }
    .row { grid-template-columns: 1fr 70px 70px 100px; font-size: 12px; }
  }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>Macro Dashboard</h1>
    <div class="sub">__SUBTITLE__</div>
  </header>

  <nav class="tabs">
    <a href="index.html" class="active">Macro</a>
    <a href="portfolio.html">Portfolio</a>
    <a href="terminal.html">Terminal</a>
    <a href="porter.html">Porter</a>
  </nav>

  <div class="calendar-label">Next 3 Weeks</div>
  <div class="calendar" id="calendar-slot">__CALENDAR__</div>

  <h2 style="margin-top:0">Charts</h2>
  <div class="controls">
    <div>
      <span class="btn-group-label">Range</span>
      <div class="btn-group" id="range-group">
        <button data-range="3m">3M</button>
        <button data-range="ytd">YTD</button>
        <button data-range="1y">1Y</button>
        <button data-range="3y">3Y</button>
        <button data-range="all" class="active">All</button>
      </div>
    </div>
    <div>
      <span class="btn-group-label">Value</span>
      <div class="btn-group" id="mode-group">
        <button data-mode="level" class="active">Level</button>
        <button data-mode="yoy">YoY %</button>
      </div>
    </div>
  </div>
  <div class="panels" id="panels"></div>

  <h2>Status Board</h2>
  <div class="board" id="board"></div>

  <footer id="footer">Source: Federal Reserve Economic Data (FRED), stlouisfed.org — generated __GENDATE__</footer>
</div>

<script>
const SUMMARY = __SUMMARY__;
const PANELS = __PANELS__;

function statusClass(status) {
  if (status === "Improving") return "up";
  if (status === "Worsening") return "down";
  return "flat";
}

function buildBoard() {
  const el = document.getElementById("board");
  const byCat = {};
  SUMMARY.forEach(r => { (byCat[r.category] = byCat[r.category] || []).push(r); });
  Object.keys(byCat).forEach(cat => {
    const group = document.createElement("div");
    group.className = "cat-group";
    group.innerHTML = `<div class="cat-head">${cat}</div>
      <div class="row head"><div>Indicator</div><div class="num">Latest</div><div class="num">Change</div><div></div></div>`;
    byCat[cat].forEach(r => {
      const row = document.createElement("div");
      row.className = "row";
      const latest = (typeof r.latest === "number") ? r.latest.toFixed(2) : r.latest;
      const change = (typeof r.change === "number") ? (r.change > 0 ? "+" : "") + r.change.toFixed(2) : "—";
      row.innerHTML = `<div class="name">${r.name} <span style="color:rgba(244,241,232,0.35)">${r.unit || ""}</span></div>
        <div class="num">${latest}</div>
        <div class="num">${change}</div>
        <div class="pill ${statusClass(r.status)}">${r.status}</div>`;
      group.appendChild(row);
    });
    el.appendChild(group);
  });
}

const COLORS = ["#d9a441", "#c96a4a", "#6a9e7f", "#8aa0c9", "#b07cc6"];
const MONTHS_BY_RANGE = { "3m": 3, "1y": 12, "3y": 36, "all": Infinity };

let currentRange = "all";
let currentMode = "level";
let chartInstances = [];

// Values are monthly, so a 12-step lookback is a year regardless of range.
function toYoY(values) {
  return values.map((v, i) => {
    const prior = values[i - 12];
    if (v === null || v === undefined || prior === null || prior === undefined || prior === 0) return null;
    return ((v - prior) / Math.abs(prior)) * 100;
  });
}

function sliceByRange(dates, seriesArr) {
  let n;
  if (currentRange === "ytd") {
    // "This year" is defined by the most recent data point's year, not the
    // viewer's real-world clock, so YTD stays correct as new data arrives.
    const lastYear = dates[dates.length - 1].split(" ")[1];
    const janIdx = dates.findIndex(d => d.startsWith("Jan") && d.endsWith(lastYear));
    n = janIdx >= 0 ? dates.length - janIdx : dates.length;
  } else {
    n = MONTHS_BY_RANGE[currentRange];
  }
  if (n === Infinity || dates.length <= n) return { dates, seriesArr };
  return {
    dates: dates.slice(-n),
    seriesArr: seriesArr.map(s => ({ ...s, values: s.values.slice(-n) })),
  };
}

function buildPanels() {
  const el = document.getElementById("panels");
  el.innerHTML = "";
  chartInstances.forEach(c => c.destroy());
  chartInstances = [];

  PANELS.forEach((p, i) => {
    const panelEl = document.createElement("div");
    panelEl.className = "panel";
    const canvasId = "chart-" + i;
    panelEl.innerHTML = `<h3>${p.title}</h3><canvas id="${canvasId}"></canvas>`;
    el.appendChild(panelEl);

    // Apply YoY transform on the FULL series first (needs 12mo of history
    // before the visible window), then slice to the selected range.
    let seriesArr = p.series;
    if (currentMode === "yoy") {
      seriesArr = seriesArr.map(s => ({ name: s.name, values: toYoY(s.values) }));
    }
    const sliced = sliceByRange(p.dates, seriesArr);

    const ctx = document.getElementById(canvasId).getContext("2d");
    const chart = new Chart(ctx, {
      type: "line",
      data: {
        labels: sliced.dates,
        datasets: sliced.seriesArr.map((s, idx) => ({
          label: s.name,
          data: s.values,
          borderColor: COLORS[idx % COLORS.length],
          backgroundColor: "transparent",
          borderWidth: 2,
          pointRadius: 0,
          pointHoverRadius: 5,
          pointHoverBackgroundColor: COLORS[idx % COLORS.length],
          pointHoverBorderColor: "#0f1a17",
          pointHoverBorderWidth: 2,
          tension: 0.15,
          spanGaps: true,
        })),
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { mode: "index", intersect: false, axis: "x" },
        plugins: {
          legend: { display: sliced.seriesArr.length > 1, labels: { color: "rgba(244,241,232,0.6)", font: { size: 10 } } },
          tooltip: {
            mode: "index",
            intersect: false,
            backgroundColor: "#1c2a25",
            titleColor: "#d9a441",
            bodyColor: "#f4f1e8",
            borderColor: "rgba(244,241,232,0.18)",
            borderWidth: 1,
            padding: 10,
            titleFont: { family: "'IBM Plex Mono', monospace", size: 11 },
            bodyFont: { family: "'IBM Plex Mono', monospace", size: 11 },
            callbacks: {
              label: ctx => {
                const v = ctx.parsed.y;
                const suffix = currentMode === "yoy" ? "%" : "";
                return ` ${ctx.dataset.label}: ${v === null || v === undefined ? "—" : v.toFixed(2) + suffix}`;
              },
            },
          },
        },
        scales: {
          x: { ticks: { color: "rgba(244,241,232,0.4)", maxTicksLimit: 8 }, grid: { color: "rgba(244,241,232,0.06)" } },
          y: {
            ticks: {
              color: "rgba(244,241,232,0.4)",
              callback: v => currentMode === "yoy" ? v + "%" : v,
            },
            grid: { color: "rgba(244,241,232,0.06)" },
          },
        },
      },
    });
    chartInstances.push(chart);
  });
}

document.getElementById("range-group").addEventListener("click", e => {
  const btn = e.target.closest("button[data-range]");
  if (!btn) return;
  currentRange = btn.dataset.range;
  document.querySelectorAll("#range-group button").forEach(b => b.classList.toggle("active", b === btn));
  buildPanels();
});

document.getElementById("mode-group").addEventListener("click", e => {
  const btn = e.target.closest("button[data-mode]");
  if (!btn) return;
  currentMode = btn.dataset.mode;
  document.querySelectorAll("#mode-group button").forEach(b => b.classList.toggle("active", b === btn));
  buildPanels();
});

buildBoard();
buildPanels();
</script>
</body>
</html>
"""


def build_html(summary, series, subtitle):
    panels_payload = []
    for title, refs in PANELS:
        s0 = series.get(refs[0])
        dates = s0["dates"] if s0 else []
        series_payload = []
        for sheet, name in refs:
            s = series.get((sheet, name))
            if not s:
                continue
            vals = s["values"]
            # pad/truncate to common date length defensively
            vals = (vals + [None] * len(dates))[:len(dates)]
            series_payload.append({"name": name, "values": vals})
        panels_payload.append({"title": title, "dates": dates, "series": series_payload})

    html = HTML_TEMPLATE.replace("__SUMMARY__", json.dumps(summary))
    html = html.replace("__PANELS__", json.dumps(panels_payload))
    html = html.replace("__SUBTITLE__", subtitle)
    html = html.replace("__CALENDAR__", build_calendar_html())
    html = html.replace("__GENDATE__", datetime.today().strftime("%Y-%m-%d"))
    return html


def main():
    ap = argparse.ArgumentParser(description="Turn macro_dashboard.xlsx into an interactive HTML dashboard.")
    ap.add_argument("workbook")
    ap.add_argument("-o", "--output", default="macro_dashboard.html")
    args = ap.parse_args()

    summary, series = read_workbook(args.workbook)
    n = len(summary)
    subtitle = f"{n} FRED indicators — Leading, Business Activity, Coincident & Lagging"
    html = build_html(summary, series, subtitle)

    with open(args.output, "w") as f:
        f.write(html)
    print(f"Wrote {args.output} ({n} indicators, {len(PANELS)} chart panels)")


if __name__ == "__main__":
    main()
