#!/bin/bash
# trigger_workflow.sh - Triggers the GitHub Actions "Update Stock Data" workflow
# at 6:00 PM IST each weekday via the repository_dispatch API (bypasses the
# codespace GITHUB_TOKEN's lack of workflow_dispatch scope). Falls back to a
# provided WORKFLOW_PAT if set. Supports IMMEDIATE=1 for a single on-demand run.

REPO="sushilade/marketcapdetails"
BRANCH="main"
EVENT_TYPE="update-stock-data"

# PAT (preferred) or the ambient codespace GITHUB_TOKEN
PAT="${WORKFLOW_PAT:-${GITHUB_TOKEN}}"
if [ -z "$PAT" ]; then
    echo "Error: No token available (set WORKFLOW_PAT, or ensure GITHUB_TOKEN is present)"
    exit 1
fi

echo "Starting workflow trigger service..."
echo "Will trigger workflow at 6:00 PM IST (12:30 PM UTC) Monday-Friday"
echo "Press Ctrl+C to stop"

trigger_workflow() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Triggering workflow via repository_dispatch..."

    response=$(curl -s -w "\n%{http_code}" \
        --request POST \
        --url "https://api.github.com/repos/${REPO}/dispatches" \
        --header "Accept: application/vnd.github+json" \
        --header "Authorization: Bearer ${PAT}" \
        --header "X-GitHub-Api-Version: 2022-11-28" \
        --data "{\"event_type\":\"${EVENT_TYPE}\",\"client_payload\":{\"ref\":\"${BRANCH}\"}}")

    http_code=$(echo "$response" | tail -n1)

    if [ "$http_code" = "204" ]; then
        echo "[$(date '+%Y-%m-%d %H:%M:%S')] Workflow triggered successfully"
    else
        echo "[$(date '+%Y-%m-%d %H:%M:%S')] Error (HTTP $http_code)"
    fi
}

# Convert 6:00 PM IST to minutes since midnight UTC
# 6:00 PM IST = 12:30 PM UTC = 750 minutes since midnight UTC
TARGET_HOUR=12
TARGET_MINUTE=30

# Immediate (one-shot) mode: trigger once and exit
if [ "${IMMEDIATE:-0}" = "1" ]; then
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Triggering workflow immediately (run-once mode)..."
    trigger_workflow
    exit 0
fi

while true; do
    now_hour=$(date -u +%-H)
    now_minute=$(date -u +%-M)
    now_day=$(date -u +%-u)  # 1=Monday, 7=Sunday

    # Check if it's a weekday (Mon-Fri)
    if [ "$now_day" -ge 1 ] && [ "$now_day" -le 5 ]; then
        # Convert current time to minutes since midnight UTC
        now_total=$(( now_hour * 60 + now_minute ))
        target_total=$(( TARGET_HOUR * 60 + TARGET_MINUTE ))

        # Check if current time is within 1 minute of target (6:00 PM IST = 12:30 PM UTC)
        if [ "$now_total" -ge "$(( target_total - 1 ))" ] && [ "$now_total" -le "$(( target_total + 1 ))" ]; then
            trigger_workflow
            # Wait 2 minutes to avoid double triggering
            sleep 120
        fi
    fi

    # Check every 30 seconds
    sleep 30
done