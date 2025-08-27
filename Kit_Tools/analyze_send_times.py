import json
import pandas as pd
from datetime import datetime
import pytz

# --- CONFIGURATION ---
BROADCASTS_FILE = 'all_broadcasts_master.json'
STATS_FILE = 'stats/all_broadcast_stats.json'
OUTPUT_FILE = 'send_time_analysis.json'
TIMEZONE = 'America/Los_Angeles' # Convert UTC to user's timezone
MIN_RECIPIENTS = 1000

# --- DATA LOADING ---
def load_data():
    """Loads broadcast and stats data from JSON files."""
    with open(BROADCASTS_FILE, 'r') as f:
        broadcasts = json.load(f)
    with open(STATS_FILE, 'r') as f:
        stats = json.load(f)
    return broadcasts, stats

# --- DATA PROCESSING ---
def process_data(broadcasts, stats):
    """Merges, filters, and prepares data for time analysis."""
    df_broadcasts = pd.DataFrame(broadcasts)
    # The stats file is a list of dicts, each with an 'id' key for the broadcast
    # Correctly flatten the deeply nested stats data from the JSON structure
    stats_list = []
    for item in stats:
        broadcast_data = item.get('broadcast', {})
        if broadcast_data:
            flat_stats = broadcast_data.get('stats', {})
            flat_stats['broadcast_id'] = broadcast_data.get('id')
            stats_list.append(flat_stats)
    df_stats = pd.DataFrame(stats_list)

    # Ensure broadcast_id is int for merging
    df_broadcasts['id'] = df_broadcasts['id'].astype(int)
    df_stats['broadcast_id'] = df_stats['broadcast_id'].astype(int)

    # Merge dataframes
    df = pd.merge(df_broadcasts, df_stats, left_on='id', right_on='broadcast_id', how='inner')

    # Filter for sent, full-list broadcasts using the 'status' field from the stats data
    df = df[df['status'] == 'completed']
    df = df[df['recipients'] >= MIN_RECIPIENTS]

    # Convert timestamps and extract time features
    df['published_at_dt'] = pd.to_datetime(df['published_at'])
    df['published_at_local'] = df['published_at_dt'].dt.tz_convert(TIMEZONE)
    df['day_of_week'] = df['published_at_local'].dt.day_name()
    df['hour_of_day'] = df['published_at_local'].dt.hour

    # Calculate open and click rates using the correct field names
    df['open_rate'] = (df['emails_opened'] / df['recipients']) * 100
    df['click_rate'] = (df['total_clicks'] / df['recipients']) * 100

    return df

# --- ANALYSIS ---
def analyze_by_time(df):
    """Analyzes performance by day of week and hour of day."""
    # Analysis by Day of Week
    day_analysis = df.groupby('day_of_week').agg(
        broadcast_count=('id', 'count'),
        avg_open_rate=('open_rate', 'mean'),
        avg_click_rate=('click_rate', 'mean')
    ).round(2).reindex(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'])

    # Analysis by Hour of Day
    hour_analysis = df.groupby('hour_of_day').agg(
        broadcast_count=('id', 'count'),
        avg_open_rate=('open_rate', 'mean'),
        avg_click_rate=('click_rate', 'mean')
    ).round(2)

    return day_analysis.to_dict('index'), hour_analysis.to_dict('index')

# --- MAIN EXECUTION ---
def main():
    """Main function to run the analysis and save results."""
    print("Starting send time analysis...")
    broadcasts, stats = load_data()
    df = process_data(broadcasts, stats)
    print(f"Analyzing {len(df)} full-list broadcasts for send time patterns...")

    day_results, hour_results = analyze_by_time(df)

    results = {
        'day_of_week_analysis': day_results,
        'hour_of_day_analysis': hour_results,
        'total_broadcasts_analyzed': len(df)
    }

    with open(OUTPUT_FILE, 'w') as f:
        json.dump(results, f, indent=2)

    print(f"Send time analysis complete. Results saved to {OUTPUT_FILE}")

if __name__ == '__main__':
    main()
