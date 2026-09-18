"""
Insights generator
====================
Flags "drastic" moves in your macro/portfolio data (>=5% weekly move on
a macro indicator, >=8% weekly move on a position) and uses the Claude
API — with web search — to write a short, real explanation for each one.

THIS COSTS REAL MONEY, unlike every other script in this pipeline. Web
search is $10 per 1,000 searches plus normal token costs on top. To keep
that low:
  - Only flagged (drastic) items get queried at all.
  - Results are cached in insights_cache.json. A flagged item is only
    re-queried if its magnitude has moved meaningfully since the last
    explanation was generated — not on every 2-hour refresh.
  - Uses Haiku (the cheapest model), with the older, broadly-supported
    web search tool version (web_search_20250305) since the newer
    dynamic-filtering version isn't available on Haiku.

SETUP
-----
1. Get an API key at https://console.anthropic.com/ (pay-as-you-go,
   separate from a Claude.ai subscription).
2. Add ANTHROPIC_API_KEY=your_key to your .env file.

USAGE
-----
python generate_insights.py macro_dashboard.xlsx --xlsx KWP.xlsx
python generate_insights.py macro_dashboard.xlsx --sheet-csv-url "https://..."

Writes insights.json. terminal_visualize.py reads it automatically if
it's sitting in the same folder — no flag needed there.
"""

import argparse
import json
import os
import urllib.request
from datetime import datetime

import macro_visualize as mv
import portfolio_visualize as pv
import terminal_visualize as tv

CACHE_PATH = "insights_cache.json"
API_URL = "https://api.anthropic.com/v1/messages"
MODEL = "claude-haiku-4-5-20251001"
WEB_SEARCH_TOOL = {"type": "web_search_20250305", "name": "web_search", "max_uses": 3}

MACRO_THRESHOLD_PCT = 5.0      # flag a macro indicator if |1w %change| >= this
PORTFOLIO_THRESHOLD_PCT = 8.0  # flag a position if |1w %change| >= this
DAILY_THRESHOLD_PCT = 4.0      # flag a position if |1d %change| >= this (needs the changepct sheet column)
RECHANGE_TOLERANCE = 0.25      # re-query only if magnitude shifted >25% (relative) since cached


def load_cache():
    if os.path.exists(CACHE_PATH):
        with open(CACHE_PATH) as f:
            return json.load(f)
    return {}


def save_cache(cache):
    with open(CACHE_PATH, "w") as f:
        json.dump(cache, f, indent=2)


def find_drastic(macro_path, portfolio_data):
    summary, series = mv.read_workbook(macro_path)
    flagged = []
    # Macro data is monthly, so it only ever has a meaningful "weekly"
    # reading — there's no daily FRED figure to flag on.
    for refs in [tv.RATES_PANEL, tv.INFLATION_PANEL, tv.LABOR_PANEL, tv.RISK_PANEL]:
        for sheet, name in refs:
            s = series.get((sheet, name))
            if not s:
                continue
            periods = tv._macro_periods(s["dates"], s["values"])
            d = periods.get("1w")
            if d and d["pct"] is not None and abs(d["pct"]) >= MACRO_THRESHOLD_PCT:
                flagged.append({
                    "key": f"macro:{name}", "label": name, "period_label": "past week",
                    "period": "weekly", "source": "macro", "pt": d["pt"], "pct": d["pct"],
                })

    for r in pv.build_rows(portfolio_data):
        if r["price"] is None:
            continue
        if r["1w"] is not None and abs(r["1w"]) >= PORTFOLIO_THRESHOLD_PCT:
            flagged.append({
                "key": f"portfolio-weekly:{r['symbol']}", "label": r["symbol"], "period_label": "past week",
                "period": "weekly", "source": "portfolio", "pt": None, "pct": r["1w"],
            })
        if r["1d"] is not None and abs(r["1d"]) >= DAILY_THRESHOLD_PCT:
            flagged.append({
                "key": f"portfolio-daily:{r['symbol']}", "label": r["symbol"], "period_label": "today",
                "period": "daily", "source": "portfolio", "pt": None, "pct": r["1d"],
            })
    return flagged


def needs_refresh(item, cache):
    cached = cache.get(item["key"])
    if not cached or cached.get("pct") is None:
        return True
    old_pct = cached["pct"]
    if old_pct == 0:
        return item["pct"] != 0
    return abs(item["pct"] - old_pct) / abs(old_pct) > RECHANGE_TOLERANCE


def explain(api_key, item):
    move_desc = f'{item["pt"]:+.2f} ({item["pct"]:+.1f}%)' if item["pt"] is not None else f'{item["pct"]:+.1f}%'
    prompt = (
        f'{item["label"]} moved {move_desc} over the {item["period_label"]}. '
        f'Search for recent news and explain the likely driver(s) in 2-3 concise '
        f'sentences. No hedging filler — just the facts.'
    )
    body = {
        "model": MODEL, "max_tokens": 350,
        "messages": [{"role": "user", "content": prompt}],
        "tools": [WEB_SEARCH_TOOL],
    }
    req = urllib.request.Request(
        API_URL, data=json.dumps(body).encode(),
        headers={"x-api-key": api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=90) as resp:
        data = json.loads(resp.read())
    text = " ".join(b["text"] for b in data.get("content", []) if b.get("type") == "text").strip()
    return text or "No explanation generated."


def main():
    ap = argparse.ArgumentParser(description="Generate cached AI explanations for drastic moves.")
    ap.add_argument("macro_workbook")
    ap.add_argument("--sheet-csv-url", default=pv.SHEET_CSV_URL)
    ap.add_argument("--xlsx", default=None)
    ap.add_argument("--demo", action="store_true")
    ap.add_argument("--api-key", default=os.environ.get("ANTHROPIC_API_KEY"))
    ap.add_argument("-o", "--output", default="insights.json")
    args = ap.parse_args()

    if not args.api_key:
        raise SystemExit("No ANTHROPIC_API_KEY found — add it to your .env file, or pass --api-key.")

    if args.demo:
        portfolio_data = pv.demo_price_data()
    elif args.xlsx:
        portfolio_data = pv.read_xlsx_price_data(args.xlsx)
    else:
        portfolio_data = pv.fetch_price_data(args.sheet_csv_url)

    flagged = find_drastic(args.macro_workbook, portfolio_data)
    cache = load_cache()
    calls_made = 0

    for item in flagged:
        if needs_refresh(item, cache):
            print(f"Querying Claude for {item['label']} ({item['pct']:+.1f}%)...")
            text = explain(args.api_key, item)
            cache[item["key"]] = {
                "label": item["label"], "pt": item["pt"], "pct": item["pct"],
                "period_label": item["period_label"], "period": item["period"],
                "source": item["source"], "key": item["key"], "text": text,
                "generated_at": datetime.today().strftime("%Y-%m-%d %H:%M"),
            }
            calls_made += 1
        else:
            print(f"Reusing cached explanation for {item['label']} (no material change)")

    save_cache(cache)

    current_keys = {item["key"] for item in flagged}
    output = sorted((cache[k] for k in current_keys), key=lambda x: abs(x["pct"]), reverse=True)
    with open(args.output, "w") as f:
        json.dump(output, f, indent=2)

    print(f"Flagged: {len(flagged)}  API calls made: {calls_made}  Wrote: {args.output}")


if __name__ == "__main__":
    main()
