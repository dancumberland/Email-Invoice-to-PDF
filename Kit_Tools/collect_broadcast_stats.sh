#!/bin/bash

# Script to collect broadcast statistics for all broadcasts
# Rate limit: 120 requests per 60 seconds, so we'll do 100 requests per minute to be safe

API_KEY="REMOVED_KIT_KEY_REVOKE"
BROADCAST_IDS_FILE="/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Kit_Tools/all_broadcasts_master.json"
OUTPUT_DIR="/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Kit_Tools/stats"

# Create output directory
mkdir -p "$OUTPUT_DIR"

# Extract broadcast IDs
echo "Extracting broadcast IDs..."
jq -r '.[].id' "$BROADCAST_IDS_FILE" > "$OUTPUT_DIR/broadcast_ids.txt"

# Count total broadcasts
TOTAL_BROADCASTS=$(wc -l < "$OUTPUT_DIR/broadcast_ids.txt")
echo "Total broadcasts to process: $TOTAL_BROADCASTS"

# Initialize counters
PROCESSED=0
BATCH_SIZE=100
BATCH_COUNT=0

echo "Starting stats collection..."

while IFS= read -r broadcast_id; do
    echo "Processing broadcast $broadcast_id ($((PROCESSED + 1))/$TOTAL_BROADCASTS)..."
    
    # Fetch stats for this broadcast
    curl --silent --request GET \
        --url "https://api.kit.com/v4/broadcasts/$broadcast_id/stats" \
        --header "X-Kit-Api-Key: $API_KEY" \
        > "$OUTPUT_DIR/stats_$broadcast_id.json"
    
    PROCESSED=$((PROCESSED + 1))
    
    # Check if we've hit the batch limit
    if [ $((PROCESSED % BATCH_SIZE)) -eq 0 ]; then
        BATCH_COUNT=$((BATCH_COUNT + 1))
        echo "Completed batch $BATCH_COUNT ($PROCESSED broadcasts). Waiting 60 seconds for rate limit..."
        sleep 60
    else
        # Small delay between requests
        sleep 0.5
    fi
    
done < "$OUTPUT_DIR/broadcast_ids.txt"

echo "Stats collection completed! Processed $PROCESSED broadcasts."
echo "Combining all stats into master file..."

# Combine all stats files into one JSON array
echo "[" > "$OUTPUT_DIR/all_broadcast_stats.json"
first=true
for stats_file in "$OUTPUT_DIR"/stats_*.json; do
    if [ "$first" = true ]; then
        first=false
    else
        echo "," >> "$OUTPUT_DIR/all_broadcast_stats.json"
    fi
    cat "$stats_file" >> "$OUTPUT_DIR/all_broadcast_stats.json"
done
echo "]" >> "$OUTPUT_DIR/all_broadcast_stats.json"

echo "All broadcast statistics saved to: $OUTPUT_DIR/all_broadcast_stats.json"
