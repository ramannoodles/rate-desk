#!/bin/bash
# Regenerates all four pages (Macro, Portfolio, Terminal, Porter) and
# publishes them to GitHub Pages. Meant to be triggered by launchd every
# couple of hours — see com.raman.rate-desk.plist (unchanged; it just
# calls this file, so no launchd edits are needed after updating this
# script).
set -euo pipefail

# --- Config ---
PUBLISH_DIR="$HOME/rate-desk"                # the GitHub Pages repo
MACRO_XLSX="$PUBLISH_DIR/macro_dashboard.xlsx"  # macro_data_pull.py's output, in this same folder now

cd "$PUBLISH_DIR"

# --- Load API keys from a local, untracked .env file ---
if [ -f .env ]; then
  export $(grep -v '^#' .env | xargs)
fi

# --- Step 1: refresh the macro data pull (FRED -> xlsx) ---
python3 macro_data_pull.py

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
