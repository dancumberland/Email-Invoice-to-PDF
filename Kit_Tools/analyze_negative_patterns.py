import json
import pandas as pd
import re
from collections import Counter

# --- CONFIGURATION ---
BROADCASTS_FILE = 'all_broadcasts_master.json'
STATS_FILE = 'stats/all_broadcast_stats.json'
OUTPUT_FILE = 'negative_pattern_analysis.json'
MIN_RECIPIENTS = 1000
NUM_WORST_BROADCASTS = 30 # Analyze the bottom 30 broadcasts

# --- DATA LOADING & PROCESSING ---
def load_and_process_data():
    """Loads, merges, and prepares data for negative pattern analysis."""
    with open(BROADCASTS_FILE, 'r') as f:
        broadcasts = json.load(f)
    with open(STATS_FILE, 'r') as f:
        stats = json.load(f)

    df_broadcasts = pd.DataFrame(broadcasts)

    stats_list = []
    for item in stats:
        broadcast_data = item.get('broadcast', {})
        if broadcast_data:
            flat_stats = broadcast_data.get('stats', {})
            flat_stats['broadcast_id'] = broadcast_data.get('id')
            stats_list.append(flat_stats)
    df_stats = pd.DataFrame(stats_list)

    df_broadcasts['id'] = df_broadcasts['id'].astype(int)
    df_stats['broadcast_id'] = df_stats['broadcast_id'].astype(int)

    df = pd.merge(df_broadcasts, df_stats, left_on='id', right_on='broadcast_id', how='inner')

    df = df[df['status'] == 'completed']
    df = df[df['recipients'] >= MIN_RECIPIENTS]

    df['open_rate'] = (df['emails_opened'] / df['recipients']) * 100
    df['preview_text'] = df['preview_text'].fillna('')
    df['subject'] = df['subject'].fillna('')

    return df

# --- ANALYSIS FUNCTIONS ---
def get_word_counts(series):
    """Utility function to get word counts from a text series."""
    stop_words = set(['a', 'an', 'the', 'in', 'on', 'is', 'and', 'to', 'for', 'of', 'it', 'i', 'you', 'your', 'this', 'with', 'was', 'are', 'my', 's', 't', 're', 've', 'll', 'd', 'm'])
    words = ' '.join(series).lower()
    words = re.findall(r'\b\w+\b', words)
    return Counter(word for word in words if word not in stop_words)

def analyze_subject_patterns(df_worst):
    """Analyzes subject lines of the worst-performing broadcasts."""
    # Keyword Analysis
    subject_keywords = get_word_counts(df_worst['subject']).most_common(15)

    # Structural Analysis
    contains_bracket = df_worst['subject'].str.contains(r'\[|\]', regex=True).sum()
    all_caps_subjects = df_worst[df_worst['subject'].str.isupper()].shape[0]
    avg_length = df_worst['subject'].str.len().mean()

    return {
        'common_keywords': subject_keywords,
        'structural_patterns': {
            'count_with_brackets': int(contains_bracket),
            'count_all_caps': int(all_caps_subjects),
            'average_length': round(avg_length, 2)
        }
    }

def analyze_preview_text_patterns(df_worst):
    """Analyzes preview text of the worst-performing broadcasts."""
    preview_keywords = get_word_counts(df_worst['preview_text']).most_common(15)
    avg_length = df_worst['preview_text'].str.len().mean()
    
    return {
        'common_keywords': preview_keywords,
        'average_length': round(avg_length, 2)
    }

# --- MAIN EXECUTION ---
def main():
    """Main function to run all analyses and save results."""
    print("Starting negative pattern analysis...")
    df = load_and_process_data()
    
    # Isolate the worst-performing broadcasts
    df_worst = df.nsmallest(NUM_WORST_BROADCASTS, 'open_rate')
    print(f"Analyzing the {len(df_worst)} lowest-performing broadcasts...")

    subject_analysis = analyze_subject_patterns(df_worst)
    preview_analysis = analyze_preview_text_patterns(df_worst)

    results = {
        'subject_analysis': subject_analysis,
        'preview_text_analysis': preview_analysis,
        'total_broadcasts_analyzed': len(df_worst),
        'worst_broadcasts_list': df_worst[['subject', 'open_rate', 'recipients']].to_dict('records')
    }

    with open(OUTPUT_FILE, 'w') as f:
        json.dump(results, f, indent=2)

    print(f"Negative pattern analysis complete. Results saved to {OUTPUT_FILE}")

if __name__ == '__main__':
    main()
