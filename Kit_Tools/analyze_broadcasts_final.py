#!/usr/bin/env python3
"""
Kit.com Broadcast Analysis Script - Final Version
- Segments by audience size (focusing on 1000+ recipients)
- Applies time-based weighting to prioritize recent broadcasts (last 6 months)
"""

import json
from datetime import datetime, timedelta

# --- CONFIGURATION ---
TIME_WEIGHT_FACTOR = 2.0  # How much more to weigh recent broadcasts
RECENCY_THRESHOLD_MONTHS = 6
FULL_LIST_THRESHOLD = 1000 # Minimum recipients for a 'full list' broadcast

# --- SCRIPT ---

def load_data():
    """Loads all necessary JSON data files."""
    try:
        with open('all_broadcasts_master.json', 'r') as f:
            broadcasts = json.load(f)
        with open('stats/all_broadcast_stats.json', 'r') as f:
            stats = json.load(f)
        return broadcasts, stats
    except FileNotFoundError as e:
        print(f"Error: Data file not found - {e}. Please ensure data collection is complete.")
        return None, None

def merge_and_filter_data(broadcasts, stats):
    """Merges broadcast content with stats and filters for valid, sent broadcasts."""
    stats_lookup = {s['broadcast']['id']: s['broadcast']['stats'] for s in stats if 'broadcast' in s and 'id' in s['broadcast']}
    
    merged_data = []
    for b in broadcasts:
        if b['id'] in stats_lookup:
            b['performance_stats'] = stats_lookup[b['id']]
            # Filter for sent broadcasts with recipients
            if (b.get('performance_stats', {}).get('status') != 'draft' and 
                b.get('performance_stats', {}).get('recipients', 0) >= FULL_LIST_THRESHOLD and
                b.get('published_at') is not None):
                merged_data.append(b)
    return merged_data

def assign_time_weight(broadcast, recency_cutoff_date):
    """Assigns a weight based on the broadcast's publication date."""
    published_at_str = broadcast.get('published_at')
    if not published_at_str:
        return 1.0
    
    published_date = datetime.fromisoformat(published_at_str.replace('Z', '+00:00'))
    return TIME_WEIGHT_FACTOR if published_date > recency_cutoff_date else 1.0

def analyze_performance(broadcasts):
    """Analyzes performance, applying recipient and time weighting."""
    recency_cutoff_date = datetime.now().astimezone() - timedelta(days=RECENCY_THRESHOLD_MONTHS * 30)
    
    for b in broadcasts:
        recipients = b.get('performance_stats', {}).get('recipients', 0)
        time_weight = assign_time_weight(b, recency_cutoff_date)
        b['analysis_weight'] = recipients * time_weight

    # Sort by performance metrics
    top_open_performers = sorted(
        broadcasts, 
        key=lambda x: x.get('performance_stats', {}).get('open_rate', 0), 
        reverse=True
    )[:15]
    
    top_click_performers = sorted(
        broadcasts, 
        key=lambda x: x.get('performance_stats', {}).get('click_rate', 0), 
        reverse=True
    )[:15]

    # Weighted average calculations
    total_weight = sum(b['analysis_weight'] for b in broadcasts)
    weighted_opens = sum(b['analysis_weight'] * (b.get('performance_stats', {}).get('open_rate', 0) / 100) for b in broadcasts)
    weighted_clicks = sum(b['analysis_weight'] * (b.get('performance_stats', {}).get('click_rate', 0) / 100) for b in broadcasts)

    weighted_avg_open_rate = (weighted_opens / total_weight * 100) if total_weight > 0 else 0
    weighted_avg_click_rate = (weighted_clicks / total_weight * 100) if total_weight > 0 else 0

    return {
        'weighted_avg_open_rate': weighted_avg_open_rate,
        'weighted_avg_click_rate': weighted_avg_click_rate,
        'total_broadcasts_analyzed': len(broadcasts),
        'recency_cutoff_date': recency_cutoff_date.isoformat(),
        'top_open_performers': format_results(top_open_performers, 'open_rate'),
        'top_click_performers': format_results(top_click_performers, 'click_rate')
    }

def format_results(broadcast_list, metric):
    """Formats a list of broadcasts for the results JSON."""
    return [{
        'subject': b.get('subject', ''),
        'metric_value': b.get('performance_stats', {}).get(metric, 0),
        'recipients': b.get('performance_stats', {}).get('recipients', 0),
        'published_at': b.get('published_at'),
        'broadcast_id': b['id']
    } for b in broadcast_list]

def main():
    """Main execution function."""
    print("Starting final analysis with time-weighting...")
    broadcasts, stats = load_data()
    if not broadcasts or not stats:
        return

    print(f"Loaded {len(broadcasts)} total broadcasts and {len(stats)} stats entries.")
    
    # Filter for full-list broadcasts only
    full_list_broadcasts = merge_and_filter_data(broadcasts, stats)
    print(f"Found {len(full_list_broadcasts)} broadcasts sent to {FULL_LIST_THRESHOLD}+ recipients.")

    if not full_list_broadcasts:
        print("No full-list broadcasts found to analyze.")
        return

    # Perform the final, weighted analysis
    analysis_results = analyze_performance(full_list_broadcasts)

    # Save results to a new file
    output_filename = 'final_weighted_analysis_results.json'
    with open(output_filename, 'w') as f:
        json.dump(analysis_results, f, indent=2)

    print("\n=== FINAL ANALYSIS COMPLETE ===")
    print(f"Analysis focused on {analysis_results['total_broadcasts_analyzed']} full-list broadcasts.")
    print(f"Broadcasts from after {datetime.fromisoformat(analysis_results['recency_cutoff_date']).strftime('%Y-%m-%d')} were weighted {TIME_WEIGHT_FACTOR}x more heavily.")
    print(f"\nTime-Weighted Average Open Rate: {analysis_results['weighted_avg_open_rate']:.1f}%" )
    print(f"Time-Weighted Average Click Rate: {analysis_results['weighted_avg_click_rate']:.1f}%" )
    print(f"\nTop Open Performer: '{analysis_results['top_open_performers'][0]['subject']}'")
    print(f"Top Click Performer: '{analysis_results['top_click_performers'][0]['subject']}'")
    print(f"\nResults saved to {output_filename}")

if __name__ == "__main__":
    main()
