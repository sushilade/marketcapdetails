#!/usr/bin/env python3
"""Check if today is an Indian market holiday."""

import json
import sys
from datetime import date
from pathlib import Path

HOLIDAY_FILE = Path(__file__).parent / "nse_holidays_2026.json"


def load_holidays():
    if not HOLIDAY_FILE.exists():
        print(f"Warning: {HOLIDAY_FILE} not found, assuming no holidays", file=sys.stderr)
        return []
    with open(HOLIDAY_FILE) as f:
        data = json.load(f)
    return data.get("holidays", [])


def is_market_holiday(today: date):
    holidays = load_holidays()
    today_str = today.isoformat()
    for h in holidays:
        if h["date"] == today_str:
            return True
    return False


if __name__ == "__main__":
    today = date.today()
    if is_market_holiday(today):
        print("HOLIDAY")
    else:
        print("WORKING")
