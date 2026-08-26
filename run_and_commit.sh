#!/bin/bash
# run_and_commit.sh - Runs apinse.py, updates index.html via web.py, and auto-commits to GitHub

cd /workspaces/marketcapdetails

echo "Running apinse.py..."
python apinse.py

echo ""
echo "Running web.py to update index.html..."
python web.py

echo ""
echo "Committing and pushing to GitHub..."
git add -A
git commit -m "Update stock data and index.html: $(date '+%Y-%m-%d %H:%M')"
git push origin main

echo ""
echo "Done! Changes saved to GitHub."
