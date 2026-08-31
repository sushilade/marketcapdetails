#!/bin/bash
# trigger_workflow.sh - Triggers the GitHub Actions workflow every 10 minutes

REPO="sushilade/marketcapdetails"
WORKFLOW_ID="343063003"
BRANCH="main"

# Get PAT from environment variable
PAT="${WORKFLOW_PAT}"

if [ -z "$PAT" ]; then
    echo "Error: WORKFLOW_PAT environment variable not set"
    echo "Set it with: export WORKFLOW_PAT=your_token_here"
    exit 1
fi

echo "Starting workflow trigger service..."
echo "Will trigger workflow every 10 minutes"
echo "Press Ctrl+C to stop"

while true; do
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Triggering workflow..."

    # Trigger workflow via GitHub API using PAT
    response=$(curl -s -w "\n%{http_code}" \
        --request POST \
        --url "https://api.github.com/repos/${REPO}/actions/workflows/${WORKFLOW_ID}/dispatches" \
        --header "Accept: application/vnd.github+json" \
        --header "Authorization: Bearer ${PAT}" \
        --header "X-GitHub-Api-Version: 2022-11-28" \
        --data "{\"ref\":\"${BRANCH}\"}")

    http_code=$(echo "$response" | tail -n1)
    body=$(echo "$response" | sed '$d')

    if [ "$http_code" = "204" ]; then
        echo "[$(date '+%Y-%m-%d %H:%M:%S')] Workflow triggered successfully"
    else
        echo "[$(date '+%Y-%m-%d %H:%M:%S')] Error (HTTP $http_code): $body"
    fi

    # Wait 10 minutes
    sleep 600
done
