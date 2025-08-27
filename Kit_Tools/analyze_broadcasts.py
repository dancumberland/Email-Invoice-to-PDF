#!/usr/bin/env python3
"""
Kit.com Broadcast Analysis Script
Analyzes broadcast performance data to identify optimization patterns
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

def analyze_subject_lines(broadcasts):
    """Analyze subject line patterns and their performance"""
    subject_analysis = {
        'length_performance': defaultdict(list),
        'question_vs_statement': {'questions': [], 'statements': []},
        'urgency_words': defaultdict(list),
        'personalization': {'personalized': [], 'generic': []},
        'top_performers': [],
        'patterns': defaultdict(list)
    }
    
    urgency_keywords = ['urgent', 'now', 'today', 'deadline', 'limited', 'expires', 'last chance', 'hurry', 'quick', 'fast']
    
    for broadcast in broadcasts:
        subject = broadcast.get('subject', '')
        stats = broadcast.get('performance_stats', {})
        open_rate = stats.get('open_rate', 0)
        recipients = stats.get('recipients', 0)
        
        # Skip if no meaningful data
        if recipients < 100:  # Filter out very small sends
            continue
        
        # Length analysis
        length = len(subject)
        length_bucket = f"{(length // 10) * 10}-{(length // 10) * 10 + 9}"
        subject_analysis['length_performance'][length_bucket].append(open_rate)
        
        # Question vs statement
        if subject.strip().endswith('?'):
            subject_analysis['question_vs_statement']['questions'].append(open_rate)
        else:
            subject_analysis['question_vs_statement']['statements'].append(open_rate)
        
        # Urgency analysis
        subject_lower = subject.lower()
        has_urgency = any(word in subject_lower for word in urgency_keywords)
        if has_urgency:
            subject_analysis['urgency_words']['with_urgency'].append(open_rate)
        else:
            subject_analysis['urgency_words']['without_urgency'].append(open_rate)
        
        # Personalization (check for first name placeholder)
        if '{{' in subject or 'first_name' in subject:
            subject_analysis['personalization']['personalized'].append(open_rate)
        else:
            subject_analysis['personalization']['generic'].append(open_rate)
        
        # Store for top performers analysis
        subject_analysis['top_performers'].append({
            'subject': subject,
            'open_rate': open_rate,
            'recipients': recipients,
            'broadcast_id': broadcast['id']
        })
    
    # Sort top performers
    subject_analysis['top_performers'].sort(key=lambda x: x['open_rate'], reverse=True)
    
    return subject_analysis

def analyze_content_performance(broadcasts):
    """Analyze content patterns and click-through performance"""
    content_analysis = {
        'length_performance': defaultdict(list),
        'cta_analysis': defaultdict(list),
        'link_count': defaultdict(list),
        'top_performers': [],
        'content_patterns': defaultdict(list)
    }
    
    for broadcast in broadcasts:
        content = broadcast.get('content', '')
        stats = broadcast.get('performance_stats', {})
        click_rate = stats.get('click_rate', 0)
        recipients = stats.get('recipients', 0)
        
        # Skip if no meaningful data
        if recipients < 100:
            continue
        
        # Content length analysis (approximate word count)
        word_count = len(content.split())
        length_bucket = f"{(word_count // 200) * 200}-{(word_count // 200) * 200 + 199}"
        content_analysis['length_performance'][length_bucket].append(click_rate)
        
        # Count links/CTAs in content
        link_count = content.count('href=')
        link_bucket = f"{min(link_count, 5)}+" if link_count >= 5 else str(link_count)
        content_analysis['link_count'][link_bucket].append(click_rate)
        
        # Store for top performers
        content_analysis['top_performers'].append({
            'subject': broadcast.get('subject', ''),
            'click_rate': click_rate,
            'recipients': recipients,
            'broadcast_id': broadcast['id'],
            'word_count': word_count,
            'link_count': link_count
        })
    
    # Sort top performers
    content_analysis['top_performers'].sort(key=lambda x: x['click_rate'], reverse=True)
    
    return content_analysis

def generate_insights(subject_analysis, content_analysis):
    """Generate actionable insights from the analysis"""
    insights = {
        'subject_line_insights': {},
        'content_insights': {},
        'recommendations': []
    }
    
    # Subject line insights
    if subject_analysis['question_vs_statement']['questions']:
        avg_question_open = statistics.mean(subject_analysis['question_vs_statement']['questions'])
        avg_statement_open = statistics.mean(subject_analysis['question_vs_statement']['statements'])
        
        insights['subject_line_insights']['question_vs_statement'] = {
            'questions_avg_open_rate': avg_question_open,
            'statements_avg_open_rate': avg_statement_open,
            'better_performer': 'questions' if avg_question_open > avg_statement_open else 'statements'
        }
    
    # Urgency analysis
    if subject_analysis['urgency_words']['with_urgency']:
        avg_urgency_open = statistics.mean(subject_analysis['urgency_words']['with_urgency'])
        avg_no_urgency_open = statistics.mean(subject_analysis['urgency_words']['without_urgency'])
        
        insights['subject_line_insights']['urgency_impact'] = {
            'with_urgency_avg_open_rate': avg_urgency_open,
            'without_urgency_avg_open_rate': avg_no_urgency_open,
            'urgency_helps': avg_urgency_open > avg_no_urgency_open
        }
    
    # Content insights
    # Find optimal content length
    length_performance = {}
    for length_bucket, rates in content_analysis['length_performance'].items():
        if rates:  # Only include buckets with data
            length_performance[length_bucket] = statistics.mean(rates)
    
    if length_performance:
        best_length = max(length_performance.items(), key=lambda x: x[1])
        insights['content_insights']['optimal_length'] = {
            'best_length_bucket': best_length[0],
            'avg_click_rate': best_length[1]
        }
    
    return insights

def save_analysis_results(subject_analysis, content_analysis, insights):
    """Save analysis results to JSON file"""
    results = {
        'analysis_date': datetime.now().isoformat(),
        'subject_line_analysis': subject_analysis,
        'content_analysis': content_analysis,
        'insights': insights
    }
    
    # Convert defaultdict to regular dict for JSON serialization
    def convert_defaultdict(obj):
        if isinstance(obj, defaultdict):
            return dict(obj)
        elif isinstance(obj, dict):
            return {k: convert_defaultdict(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [convert_defaultdict(item) for item in obj]
        else:
            return obj
    
    results = convert_defaultdict(results)
    
    with open('/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Kit_Tools/analysis_results.json', 'w') as f:
        json.dump(results, f, indent=2)
    
    print("Analysis results saved to analysis_results.json")

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
    
    print("Analyzing subject line patterns...")
    subject_analysis = analyze_subject_lines(sent_broadcasts)
    
    print("Analyzing content performance...")
    content_analysis = analyze_content_performance(sent_broadcasts)
    
    print("Generating insights...")
    insights = generate_insights(subject_analysis, content_analysis)
    
    print("Saving analysis results...")
    save_analysis_results(subject_analysis, content_analysis, insights)
    
    # Print summary
    print("\n=== ANALYSIS SUMMARY ===")
    print(f"Total broadcasts analyzed: {len(sent_broadcasts)}")
    
    if subject_analysis['top_performers']:
        top_subject = subject_analysis['top_performers'][0]
        print(f"Best performing subject line: '{top_subject['subject'][:50]}...' ({top_subject['open_rate']:.1%} open rate)")
    
    if content_analysis['top_performers']:
        top_content = content_analysis['top_performers'][0]
        print(f"Best performing content: '{top_content['subject'][:50]}...' ({top_content['click_rate']:.1%} click rate)")
    
    print("\nDetailed results saved to analysis_results.json")

if __name__ == "__main__":
    main()
