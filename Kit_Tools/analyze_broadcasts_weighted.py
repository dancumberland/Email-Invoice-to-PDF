#!/usr/bin/env python3
"""
Kit.com Broadcast Analysis Script - Weighted by Recipient Count
Analyzes broadcast performance with proper segmentation by audience size
"""

import json
import re
from datetime import datetime
from collections import defaultdict, Counter
import statistics

def load_broadcast_data():
    """Load broadcast content and statistics data"""
    try:
        # Load broadcast content
        with open('/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Kit_Tools/all_broadcasts_master.json', 'r') as f:
            broadcasts = json.load(f)
        
        # Load broadcast statistics
        with open('/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Kit_Tools/stats/all_broadcast_stats.json', 'r') as f:
            stats = json.load(f)
        
        return broadcasts, stats
    except FileNotFoundError as e:
        print(f"Data file not found: {e}")
        return None, None

def merge_broadcast_data(broadcasts, stats):
    """Merge broadcast content with performance statistics"""
    # Create a lookup dictionary for stats by broadcast ID
    stats_lookup = {}
    for stat_entry in stats:
        if 'broadcast' in stat_entry and 'id' in stat_entry['broadcast']:
            broadcast_id = stat_entry['broadcast']['id']
            stats_lookup[broadcast_id] = stat_entry['broadcast']['stats']
    
    # Merge data
    merged_data = []
    for broadcast in broadcasts:
        broadcast_id = broadcast['id']
        if broadcast_id in stats_lookup:
            broadcast_with_stats = broadcast.copy()
            broadcast_with_stats['performance_stats'] = stats_lookup[broadcast_id]
            merged_data.append(broadcast_with_stats)
    
    return merged_data

def filter_sent_broadcasts(merged_data):
    """Filter to only include broadcasts that were actually sent"""
    sent_broadcasts = []
    for broadcast in merged_data:
        # Check if broadcast was sent (not draft, has recipients, has open data)
        stats = broadcast.get('performance_stats', {})
        if (stats.get('status') != 'draft' and 
            stats.get('recipients', 0) > 0 and
            broadcast.get('published_at') is not None):
            sent_broadcasts.append(broadcast)
    
    return sent_broadcasts

def segment_broadcasts_by_size(broadcasts):
    """Segment broadcasts by recipient count"""
    segments = {
        'full_list': [],      # 1000+ recipients (full list broadcasts)
        'large_segment': [],  # 500-999 recipients (large segments)
        'medium_segment': [], # 200-499 recipients (medium segments)
        'small_segment': []   # <200 recipients (small, highly-engaged segments)
    }
    
    for broadcast in broadcasts:
        recipients = broadcast.get('performance_stats', {}).get('recipients', 0)
        
        if recipients >= 1000:
            segments['full_list'].append(broadcast)
        elif recipients >= 500:
            segments['large_segment'].append(broadcast)
        elif recipients >= 200:
            segments['medium_segment'].append(broadcast)
        else:
            segments['small_segment'].append(broadcast)
    
    return segments

def calculate_weighted_metrics(broadcasts):
    """Calculate performance metrics weighted by recipient count"""
    total_recipients = sum(b.get('performance_stats', {}).get('recipients', 0) for b in broadcasts)
    
    if total_recipients == 0:
        return {'weighted_open_rate': 0, 'weighted_click_rate': 0}
    
    weighted_opens = sum(
        b.get('performance_stats', {}).get('recipients', 0) * 
        b.get('performance_stats', {}).get('open_rate', 0) / 100
        for b in broadcasts
    )
    
    weighted_clicks = sum(
        b.get('performance_stats', {}).get('recipients', 0) * 
        b.get('performance_stats', {}).get('click_rate', 0) / 100
        for b in broadcasts
    )
    
    return {
        'weighted_open_rate': (weighted_opens / total_recipients) * 100,
        'weighted_click_rate': (weighted_clicks / total_recipients) * 100,
        'total_recipients': total_recipients,
        'total_broadcasts': len(broadcasts)
    }

def analyze_segment_performance(segment_broadcasts, segment_name):
    """Analyze performance for a specific segment"""
    if not segment_broadcasts:
        return None
    
    # Calculate weighted metrics
    weighted_metrics = calculate_weighted_metrics(segment_broadcasts)
    
    # Top performers by absolute performance (but within this segment)
    top_open_performers = sorted(
        segment_broadcasts, 
        key=lambda x: x.get('performance_stats', {}).get('open_rate', 0), 
        reverse=True
    )[:10]
    
    top_click_performers = sorted(
        segment_broadcasts, 
        key=lambda x: x.get('performance_stats', {}).get('click_rate', 0), 
        reverse=True
    )[:10]
    
    # Subject line analysis for this segment
    subject_analysis = analyze_subject_lines_segment(segment_broadcasts)
    
    return {
        'segment_name': segment_name,
        'metrics': weighted_metrics,
        'top_open_performers': [{
            'subject': b.get('subject', ''),
            'open_rate': b.get('performance_stats', {}).get('open_rate', 0),
            'recipients': b.get('performance_stats', {}).get('recipients', 0),
            'broadcast_id': b['id']
        } for b in top_open_performers],
        'top_click_performers': [{
            'subject': b.get('subject', ''),
            'click_rate': b.get('performance_stats', {}).get('click_rate', 0),
            'recipients': b.get('performance_stats', {}).get('recipients', 0),
            'broadcast_id': b['id']
        } for b in top_click_performers],
        'subject_analysis': subject_analysis
    }

def analyze_subject_lines_segment(broadcasts):
    """Analyze subject line patterns for a specific segment"""
    if not broadcasts:
        return {}
    
    urgency_keywords = ['urgent', 'now', 'today', 'deadline', 'limited', 'expires', 'last chance', 'hurry', 'quick', 'fast', 'closing']
    
    questions = []
    statements = []
    with_urgency = []
    without_urgency = []
    
    total_weighted_opens = 0
    total_recipients = 0
    
    for broadcast in broadcasts:
        subject = broadcast.get('subject', '')
        stats = broadcast.get('performance_stats', {})
        open_rate = stats.get('open_rate', 0)
        recipients = stats.get('recipients', 0)
        
        if recipients == 0:
            continue
        
        # Weight by recipient count
        weighted_opens = recipients * (open_rate / 100)
        total_weighted_opens += weighted_opens
        total_recipients += recipients
        
        # Question vs statement
        if subject.strip().endswith('?'):
            questions.append({'open_rate': open_rate, 'recipients': recipients, 'weighted_opens': weighted_opens})
        else:
            statements.append({'open_rate': open_rate, 'recipients': recipients, 'weighted_opens': weighted_opens})
        
        # Urgency analysis
        subject_lower = subject.lower()
        has_urgency = any(word in subject_lower for word in urgency_keywords)
        if has_urgency:
            with_urgency.append({'open_rate': open_rate, 'recipients': recipients, 'weighted_opens': weighted_opens})
        else:
            without_urgency.append({'open_rate': open_rate, 'recipients': recipients, 'weighted_opens': weighted_opens})
    
    def calculate_weighted_average(items):
        if not items:
            return 0
        total_weighted = sum(item['weighted_opens'] for item in items)
        total_recip = sum(item['recipients'] for item in items)
        return (total_weighted / total_recip * 100) if total_recip > 0 else 0
    
    return {
        'questions_weighted_avg': calculate_weighted_average(questions),
        'statements_weighted_avg': calculate_weighted_average(statements),
        'urgency_weighted_avg': calculate_weighted_average(with_urgency),
        'no_urgency_weighted_avg': calculate_weighted_average(without_urgency),
        'questions_count': len(questions),
        'statements_count': len(statements),
        'urgency_count': len(with_urgency),
        'no_urgency_count': len(without_urgency)
    }

def main():
    """Main analysis function"""
    print("Loading broadcast data...")
    broadcasts, stats = load_broadcast_data()
    
    if broadcasts is None or stats is None:
        print("Could not load data files. Make sure the data collection is complete.")
        return
    
    print(f"Loaded {len(broadcasts)} broadcasts and {len(stats)} stat entries")
    
    print("Merging broadcast content with performance data...")
    merged_data = merge_broadcast_data(broadcasts, stats)
    print(f"Merged data for {len(merged_data)} broadcasts")
    
    print("Filtering to sent broadcasts only...")
    sent_broadcasts = filter_sent_broadcasts(merged_data)
    print(f"Found {len(sent_broadcasts)} sent broadcasts for analysis")
    
    if len(sent_broadcasts) == 0:
        print("No sent broadcasts found for analysis.")
        return
    
    print("Segmenting broadcasts by recipient count...")
    segments = segment_broadcasts_by_size(sent_broadcasts)
    
    # Print segment summary
    print("\n=== SEGMENT SUMMARY ===")
    for segment_name, segment_broadcasts in segments.items():
        if segment_broadcasts:
            recipients = [b.get('performance_stats', {}).get('recipients', 0) for b in segment_broadcasts]
            print(f"{segment_name}: {len(segment_broadcasts)} broadcasts, {min(recipients)}-{max(recipients)} recipients")
    
    # Analyze each segment
    segment_analyses = {}
    for segment_name, segment_broadcasts in segments.items():
        if segment_broadcasts:
            print(f"\nAnalyzing {segment_name} segment...")
            analysis = analyze_segment_performance(segment_broadcasts, segment_name)
            segment_analyses[segment_name] = analysis
    
    # Save results
    results = {
        'analysis_date': datetime.now().isoformat(),
        'segment_analyses': segment_analyses,
        'overall_metrics': calculate_weighted_metrics(sent_broadcasts)
    }
    
    with open('/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Kit_Tools/weighted_analysis_results.json', 'w') as f:
        json.dump(results, f, indent=2)
    
    print("\n=== WEIGHTED ANALYSIS SUMMARY ===")
    
    # Focus on full list performance
    if 'full_list' in segment_analyses:
        full_list = segment_analyses['full_list']
        print(f"\n🎯 FULL LIST PERFORMANCE (1000+ recipients):")
        print(f"   Broadcasts: {full_list['metrics']['total_broadcasts']}")
        print(f"   Total Recipients: {full_list['metrics']['total_recipients']:,}")
        print(f"   Weighted Open Rate: {full_list['metrics']['weighted_open_rate']:.1f}%")
        print(f"   Weighted Click Rate: {full_list['metrics']['weighted_click_rate']:.1f}%")
        
        if full_list['top_open_performers']:
            top_open = full_list['top_open_performers'][0]
            print(f"   Best Open Rate: '{top_open['subject'][:50]}...' ({top_open['open_rate']:.1f}%, {top_open['recipients']} recipients)")
        
        if full_list['top_click_performers']:
            top_click = full_list['top_click_performers'][0]
            print(f"   Best Click Rate: '{top_click['subject'][:50]}...' ({top_click['click_rate']:.1f}%, {top_click['recipients']} recipients)")
    
    # Compare segments
    print(f"\n📊 SEGMENT COMPARISON:")
    for segment_name, analysis in segment_analyses.items():
        if analysis:
            metrics = analysis['metrics']
            print(f"   {segment_name}: {metrics['weighted_open_rate']:.1f}% open, {metrics['weighted_click_rate']:.1f}% click ({metrics['total_broadcasts']} broadcasts)")
    
    print(f"\nDetailed results saved to weighted_analysis_results.json")

if __name__ == "__main__":
    main()
