#!/bin/bash
# Regenerates all four pages (Macro, Portfolio, Terminal, Porter) and
# publishes them to GitHub Pages. Meant to be triggered by launchd every
# couple of hours — see com.raman.rate-desk.plist (unchanged; it just
# calls this file, so no launchd edits are needed after updating this
# script).
set -euo pipefail

# --- Config: update these three paths for your machine ---
PUBLISH_DIR="$HOME/rate-desk"                          # the GitHub Pages repo
MACRO_PULL_DIR="$HOME/PycharmProjects/Claude/Macro v2"  # your existing macro_data_pull.py
MACRO_XLSX="$MACRO_PULL_DIR/macro_dashboard.xlsx"       # its output workbook

cd "$PUBLISH_DIR"

# --- Load API keys from a local, untracked .env file ---
if [ -f .env ]; then
  export $(grep -v '^#' .env | xargs)
fi

# --- Step 1: refresh the macro data pull (FRED + yfinance -> xlsx) ---
cd "$MACRO_PULL_DIR"
source .venv/bin/activate
python macro_data_pull.py
deactivate
cd "$PUBLISH_DIR"

# --- Step 2: regenerate all four pages from that data ---
python3 macro_visualize.py "$MACRO_XLSX" -o index.html
python3 portfolio_visualize.py -o portfolio.html
python3 terminal_visualize.py "$MACRO_XLSX" -o terminal.html
python3 porter_visualize.py -o porter.html

# --- Step 3: publish only if something actually changed ---
git add index.html portfolio.html terminal.html porter.html

if git diff --cached --quiet; then
  echo "$(date '+%Y-%m-%d %H:%M:%S') No changes, skipping commit."
  exit 0
fi

git commit -m "Update dashboard $(date '+%Y-%m-%d %H:%M')"
git push origin main

echo "$(date '+%Y-%m-%d %H:%M:%S') Published."
