#!/usr/bin/env python3
"""
Batch Voice Document Updater for Dan Cumberland
Processes emails in batches of 100 and progressively updates the brand voice document.
"""

import json
import os
from datetime import datetime
from typing import List, Dict, Any

class BatchVoiceUpdater:
    def __init__(self, dataset_path: str, voice_doc_path: str):
        self.dataset_path = dataset_path
        self.voice_doc_path = voice_doc_path
        self.dataset = None
        self.current_batch = 0
        self.batch_size = 100
        
    def load_dataset(self):
        """Load the email dataset."""
        print("📂 Loading email dataset...")
        with open(self.dataset_path, 'r', encoding='utf-8') as f:
            self.dataset = json.load(f)
        
        total_emails = len(self.dataset.get('emails', []))
        print(f"✅ Loaded {total_emails} conversational emails")
        return total_emails
    
    def get_batch_emails(self, batch_num: int) -> List[Dict]:
        """Get emails for a specific batch."""
        emails = self.dataset.get('emails', [])
        start_idx = batch_num * self.batch_size
        end_idx = min(start_idx + self.batch_size, len(emails))
        
        batch_emails = emails[start_idx:end_idx]
        print(f"📦 Batch {batch_num + 1}: Processing emails {start_idx + 1}-{end_idx}")
        return batch_emails
    
    def analyze_batch_patterns(self, batch_emails: List[Dict]) -> Dict[str, Any]:
        """Analyze voice patterns in a batch of emails."""
        patterns = {
            'greeting_styles': {},
            'sign_off_styles': {},
            'question_counts': [],
            'exclamation_counts': [],
            'personal_reference_counts': [],
            'helping_phrase_counts': [],
            'collaborative_language_counts': [],
            'email_categories': {},
            'common_phrases': {},
            'sentence_lengths': [],
            'uses_contractions': 0,
            'total_contractions': 0
        }
        
        for email in batch_emails:
            voice_data = email.get('voice_patterns', {})
            raw_data = email.get('raw_data', {})
            content = raw_data.get('content', '')
            
            # Collect greeting styles
            greeting_style = voice_data.get('greeting_style', 'unknown')
            patterns['greeting_styles'][greeting_style] = patterns['greeting_styles'].get(greeting_style, 0) + 1
            
            # Collect sign-off styles
            sign_off_style = voice_data.get('sign_off_style', 'unknown')
            patterns['sign_off_styles'][sign_off_style] = patterns['sign_off_styles'].get(sign_off_style, 0) + 1
            
            # Collect numeric patterns
            patterns['question_counts'].append(voice_data.get('question_count', 0))
            patterns['exclamation_counts'].append(voice_data.get('exclamation_count', 0))
            patterns['personal_reference_counts'].append(voice_data.get('personal_references', 0))
            patterns['helping_phrase_counts'].append(voice_data.get('helping_phrases', 0))
            patterns['collaborative_language_counts'].append(voice_data.get('collaborative_language', 0))
            patterns['sentence_lengths'].append(voice_data.get('avg_sentence_length', 0))
            
            # Track contractions
            if voice_data.get('uses_contractions', False):
                patterns['uses_contractions'] += 1
            patterns['total_contractions'] += 1
            
            # Email categories
            category = email.get('category', 'unknown')
            patterns['email_categories'][category] = patterns['email_categories'].get(category, 0) + 1
            
            # Extract common phrases (sentences that appear frequently)
            sentences = content.replace('\r\n', ' ').replace('\n', ' ').split('.')
            for sentence in sentences:
                sentence = sentence.strip()
                if len(sentence) > 15 and len(sentence) < 80:  # Reasonable sentence length
                    # Clean up the sentence
                    sentence = sentence.replace('\r', '').replace('\n', ' ')
                    if not sentence.startswith('On ') and not sentence.startswith('>'):
                        patterns['common_phrases'][sentence] = patterns['common_phrases'].get(sentence, 0) + 1
        
        # Sort and limit results
        patterns['greeting_styles'] = dict(sorted(patterns['greeting_styles'].items(), key=lambda x: x[1], reverse=True))
        patterns['sign_off_styles'] = dict(sorted(patterns['sign_off_styles'].items(), key=lambda x: x[1], reverse=True))
        patterns['email_categories'] = dict(sorted(patterns['email_categories'].items(), key=lambda x: x[1], reverse=True))
        patterns['common_phrases'] = dict(sorted(patterns['common_phrases'].items(), key=lambda x: x[1], reverse=True)[:15])
        
        return patterns
    
    def extract_sample_emails(self, batch_emails: List[Dict], num_samples: int = 5) -> List[Dict]:
        """Extract representative sample emails from the batch."""
        samples = []
        
        # Try to get diverse samples by category
        categories_seen = set()
        for email in batch_emails:
            if len(samples) >= num_samples:
                break
                
            category = email.get('category', 'unknown')
            raw_data = email.get('raw_data', {})
            content = raw_data.get('content', '')
            
            # Skip very short emails
            if len(content) < 50:
                continue
                
            # Prefer emails from different categories
            if category not in categories_seen or len(samples) < 3:
                samples.append({
                    'subject': raw_data.get('subject', ''),
                    'content': content[:500] + '...' if len(content) > 500 else content,
                    'category': category,
                    'to': raw_data.get('to', ''),
                    'date': raw_data.get('date', '')
                })
                categories_seen.add(category)
        
        return samples
    
    def generate_batch_insights(self, batch_num: int, batch_emails: List[Dict]) -> str:
        """Generate insights and analysis for a batch of emails."""
        patterns = self.analyze_batch_patterns(batch_emails)
        samples = self.extract_sample_emails(batch_emails)
        
        # Calculate averages
        avg_questions = sum(patterns['question_counts']) / len(patterns['question_counts']) if patterns['question_counts'] else 0
        avg_exclamations = sum(patterns['exclamation_counts']) / len(patterns['exclamation_counts']) if patterns['exclamation_counts'] else 0
        avg_sentence_length = sum(patterns['sentence_lengths']) / len(patterns['sentence_lengths']) if patterns['sentence_lengths'] else 0
        contraction_rate = (patterns['uses_contractions'] / patterns['total_contractions']) * 100 if patterns['total_contractions'] > 0 else 0
        
        insights = f"""
## Batch {batch_num + 1} Analysis ({len(batch_emails)} emails)

### Voice Patterns Discovered

**Greeting Styles:**
"""
        
        for style, count in patterns['greeting_styles'].items():
            insights += f"- {style}: {count} emails\n"
        
        insights += f"""
**Sign-off Styles:**
"""
        
        for style, count in patterns['sign_off_styles'].items():
            insights += f"- {style}: {count} emails\n"
        
        insights += f"""
**Email Categories in This Batch:**
"""
        
        for category, count in patterns['email_categories'].items():
            insights += f"- {category}: {count} emails\n"
        
        insights += f"""
**Communication Style Metrics:**
- Average questions per email: {avg_questions:.1f}
- Average exclamations per email: {avg_exclamations:.1f}
- Average sentence length: {avg_sentence_length:.1f} words
- Uses contractions: {contraction_rate:.1f}% of emails

**Frequently Used Phrases:**
"""
        
        for phrase, count in list(patterns['common_phrases'].items())[:8]:
            if count > 1:  # Only show phrases used multiple times
                insights += f"- \"{phrase}\" (used {count} times)\n"
        
        insights += f"""
### Sample Emails from This Batch

"""
        
        for i, sample in enumerate(samples, 1):
            insights += f"""
**Sample {i}: {sample['category'].title()} Email**
- Subject: {sample['subject']}
- To: {sample['to']}
- Content Preview:
```
{sample['content']}
```

"""
        
        return insights
    
    def process_batch(self, batch_num: int) -> str:
        """Process a single batch and return insights."""
        batch_emails = self.get_batch_emails(batch_num)
        
        if not batch_emails:
            return None
        
        print(f"🔍 Analyzing voice patterns in batch {batch_num + 1}...")
        insights = self.generate_batch_insights(batch_num, batch_emails)
        
        print(f"✅ Batch {batch_num + 1} analysis complete")
        return insights

def main():
    """Main execution function."""
    print("Dan Cumberland Batch Voice Document Updater")
    print("=" * 60)
    
    dataset_path = "/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Gmail_Voice_Doc/email_voice_dataset_1000.json"
    voice_doc_path = "/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Gmail_Voice_Doc/Dan_Email_Voice_Training_MASTER.md"
    
    updater = BatchVoiceUpdater(dataset_path, voice_doc_path)
    
    # Load dataset
    total_emails = updater.load_dataset()
    total_batches = (total_emails + updater.batch_size - 1) // updater.batch_size
    
    print(f"📊 Total emails: {total_emails}")
    print(f"📦 Total batches: {total_batches}")
    print(f"📏 Batch size: {updater.batch_size}")
    
    # Process first batch as example
    print(f"\n🚀 Processing Batch 1 as example...")
    batch_insights = updater.process_batch(0)
    
    if batch_insights:
        print("\n" + "="*60)
        print("BATCH 1 INSIGHTS:")
        print("="*60)
        print(batch_insights)
        
        # Save batch insights to file
        insights_path = "/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Gmail_Voice_Doc/batch_1_insights.md"
        with open(insights_path, 'w', encoding='utf-8') as f:
            f.write(batch_insights)
        
        print(f"\n💾 Batch 1 insights saved to: {insights_path}")
        print(f"\n✅ Ready to update voice training document with Batch 1 insights!")
        print(f"📋 Next: Review and integrate these insights into the master voice document.")

if __name__ == "__main__":
    main()
