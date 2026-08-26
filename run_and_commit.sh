#!/bin/bash
# run_and_commit.sh - Runs apinse.py and auto-commits changes to GitHub

cd /workspaces/marketcapdetails

echo "Running apinse.py..."
python apinse.py

echo ""
echo "Committing and pushing to GitHub..."
git add -A
git commit -m "Update stock data: $(date '+%Y-%m-%d %H:%M')"
git push origin main

echo ""
echo "Done! Changes saved to GitHub."
