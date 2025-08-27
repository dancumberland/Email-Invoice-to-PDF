import json
import pandas as pd
import re
from collections import Counter

# --- CONFIGURATION ---
BROADCASTS_FILE = 'all_broadcasts_master.json'
STATS_FILE = 'stats/all_broadcast_stats.json'
OUTPUT_FILE = 'preview_text_analysis.json'
MIN_RECIPIENTS = 1000

# --- DATA LOADING & PROCESSING (from previous successful scripts) ---
def load_and_process_data():
    """Loads, merges, and prepares data for preview text analysis."""
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
    df['preview_text'] = df['preview_text'].fillna('') # Handle missing preview text

    return df

# --- ANALYSIS FUNCTIONS ---
def analyze_length(df):
    """Analyzes open rate by preview text length."""
    df['preview_length'] = df['preview_text'].str.len()
    bins = [0, 25, 50, 75, 100, 150, 1000]
    labels = ['0-25', '26-50', '51-75', '76-100', '101-150', '150+']
    df['length_bucket'] = pd.cut(df['preview_length'], bins=bins, labels=labels, right=False)

    length_analysis = df.groupby('length_bucket').agg(
        broadcast_count=('id', 'count'),
        avg_open_rate=('open_rate', 'mean')
    ).round(2)

    return length_analysis.to_dict('index')

def analyze_questions(df):
    """Analyzes open rate based on presence of a question mark."""
    df['has_question'] = df['preview_text'].str.contains('?', regex=False)
    question_analysis = df.groupby('has_question').agg(
        broadcast_count=('id', 'count'),
        avg_open_rate=('open_rate', 'mean')
    ).round(2)

    return question_analysis.to_dict('index')

def analyze_keywords(df):
    """Finds common keywords in top and bottom performing preview texts."""
    # Define top/bottom performers
    top_performers = df.nlargest(20, 'open_rate')
    bottom_performers = df.nsmallest(20, 'open_rate')

    # Simple stop words list
    stop_words = set(['a', 'an', 'the', 'in', 'on', 'is', 'and', 'to', 'for', 'of', 'it', 'i', 'you', 'your', 'this'])

    def get_word_counts(series):
        words = ' '.join(series).lower()
        words = re.findall(r'\b\w+\b', words)
        return Counter(word for word in words if word not in stop_words)

    top_keywords = get_word_counts(top_performers['preview_text']).most_common(10)
    bottom_keywords = get_word_counts(bottom_performers['preview_text']).most_common(10)

    return {
        'top_keywords': top_keywords,
        'bottom_keywords': bottom_keywords
    }

# --- MAIN EXECUTION ---
def main():
    """Main function to run all analyses and save results."""
    print("Starting preview text analysis...")
    df = load_and_process_data()
    print(f"Analyzing preview text for {len(df)} full-list broadcasts...")

    length_results = analyze_length(df)
    question_results = analyze_questions(df)
    keyword_results = analyze_keywords(df)

    results = {
        'length_analysis': length_results,
        'question_analysis': question_results,
        'keyword_analysis': keyword_results,
        'total_broadcasts_analyzed': len(df)
    }

    with open(OUTPUT_FILE, 'w') as f:
        json.dump(results, f, indent=2)

    print(f"Preview text analysis complete. Results saved to {OUTPUT_FILE}")

if __name__ == '__main__':
    main()
