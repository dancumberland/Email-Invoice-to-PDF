#!/usr/bin/env python3
"""
Comprehensive Email Voice Extractor for Dan Cumberland
Extracts all emails from the last 6 months, filters out repetitive/templated content,
and creates a structured dataset for voice training analysis.
"""

import json
import re
from datetime import datetime, timedelta
from typing import Dict, List, Set, Any
from collections import defaultdict
import hashlib

class EmailVoiceExtractor:
    def __init__(self):
        self.emails = []
        self.filtered_emails = []
        self.stats = {
            'total_emails': 0,
            'conversational_emails': 0,
            'filtered_out': 0,
            'recipients': defaultdict(int),
            'subject_patterns': defaultdict(int),
            'date_range': {'start': None, 'end': None}
        }
        
        # Patterns to identify repetitive/templated emails
        self.repetitive_patterns = [
            r"thanks for downloading.*guide",
            r"webinar access fixed",
            r"the founder's escape plan.*webinar",
            r"checking in.*this is dan from the meaning movement",
            r"your ai mistakes guide",
            r"lead score",
            r"automated.*response",
            r"unsubscribe",
            r"newsletter.*confirmation",
            r"calendar.*reminder",
            r"booking.*confirmation"
        ]
        
        # Minimum content length for conversational emails
        self.min_content_length = 50
        
        # Maximum similarity threshold for duplicate detection
        self.similarity_threshold = 0.85

    def is_repetitive_email(self, subject: str, content: str) -> bool:
        """Check if email matches repetitive/templated patterns."""
        subject_lower = subject.lower()
        content_lower = content.lower()
        
        # Check against known repetitive patterns
        for pattern in self.repetitive_patterns:
            if re.search(pattern, subject_lower, re.IGNORECASE) or \
               re.search(pattern, content_lower, re.IGNORECASE):
                return True
        
        # Check for very short emails (likely automated)
        if len(content.strip()) < self.min_content_length:
            return True
            
        # Check for common automated phrases
        automated_phrases = [
            "this is an automated",
            "do not reply to this email",
            "automatically generated",
            "system notification"
        ]
        
        for phrase in automated_phrases:
            if phrase in content_lower:
                return True
                
        return False

    def calculate_content_similarity(self, content1: str, content2: str) -> float:
        """Calculate similarity between two email contents using simple hash comparison."""
        # Normalize content (remove whitespace, convert to lowercase)
        norm1 = re.sub(r'\s+', ' ', content1.lower().strip())
        norm2 = re.sub(r'\s+', ' ', content2.lower().strip())
        
        # If contents are identical, return 1.0
        if norm1 == norm2:
            return 1.0
            
        # Simple similarity based on common words
        words1 = set(norm1.split())
        words2 = set(norm2.split())
        
        if not words1 or not words2:
            return 0.0
            
        intersection = words1.intersection(words2)
        union = words1.union(words2)
        
        return len(intersection) / len(union) if union else 0.0

    def is_duplicate_content(self, new_content: str, existing_contents: List[str]) -> bool:
        """Check if new content is too similar to existing content."""
        for existing in existing_contents:
            similarity = self.calculate_content_similarity(new_content, existing)
            if similarity > self.similarity_threshold:
                return True
        return False

    def extract_email_metadata(self, email_data: Dict) -> Dict:
        """Extract relevant metadata from email."""
        return {
            'id': email_data.get('id', ''),
            'thread_id': email_data.get('thread_id', ''),
            'subject': email_data.get('subject', ''),
            'from': email_data.get('from', ''),
            'to': email_data.get('to', ''),
            'date': email_data.get('date', ''),
            'content_length': len(email_data.get('content', '')),
            'has_attachments': bool(email_data.get('attachments', [])),
            'is_reply': email_data.get('subject', '').startswith('Re:'),
            'is_forward': email_data.get('subject', '').startswith('Fwd:')
        }

    def categorize_email(self, email_data: Dict) -> str:
        """Categorize email based on content and context."""
        subject = email_data.get('subject', '').lower()
        content = email_data.get('content', '').lower()
        to_field = email_data.get('to', '').lower()
        
        # Business categories
        if any(word in content for word in ['proposal', 'contract', 'agreement', 'pricing']):
            return 'business_formal'
        elif any(word in content for word in ['thanks', 'thank you', 'appreciate']):
            return 'gratitude'
        elif any(word in subject for word in ['meeting', 'call', 'schedule', 'calendar']):
            return 'scheduling'
        elif any(word in content for word in ['sorry', 'apologize', 'mistake', 'error']):
            return 'problem_solving'
        elif 'intro' in subject or 'introduction' in content:
            return 'introduction'
        elif email_data.get('subject', '').startswith('Re:'):
            return 'response'
        elif any(word in content for word in ['help', 'assist', 'support', 'question']):
            return 'helpful_inquiry'
        else:
            return 'general_conversation'

    def analyze_voice_patterns(self, email_data: Dict) -> Dict:
        """Analyze voice patterns in the email content."""
        content = email_data.get('content', '')
        
        patterns = {
            'greeting_style': self.extract_greeting(content),
            'sign_off_style': self.extract_sign_off(content),
            'question_count': len(re.findall(r'\?', content)),
            'exclamation_count': len(re.findall(r'!', content)),
            'personal_references': self.count_personal_references(content),
            'helping_phrases': self.count_helping_phrases(content),
            'collaborative_language': self.count_collaborative_language(content),
            'paragraph_count': len([p for p in content.split('\n\n') if p.strip()]),
            'avg_sentence_length': self.calculate_avg_sentence_length(content),
            'uses_contractions': bool(re.search(r"\b\w+'[a-z]+\b", content, re.IGNORECASE))
        }
        
        return patterns

    def extract_greeting(self, content: str) -> str:
        """Extract greeting style from email content."""
        lines = content.split('\n')
        first_line = lines[0].strip() if lines else ''
        
        greeting_patterns = [
            (r'^hey\s+\w+[!,]?', 'hey_name'),
            (r'^hi\s+\w+[!,]?', 'hi_name'),
            (r'^hello\s+\w+[!,]?', 'hello_name'),
            (r'^hey[!,]?$', 'hey_casual'),
            (r'^hi[!,]?$', 'hi_casual'),
            (r'^hello[!,]?$', 'hello_formal')
        ]
        
        for pattern, style in greeting_patterns:
            if re.search(pattern, first_line, re.IGNORECASE):
                return style
                
        return 'other'

    def extract_sign_off(self, content: str) -> str:
        """Extract sign-off style from email content."""
        lines = content.split('\n')
        last_lines = lines[-3:] if len(lines) >= 3 else lines
        
        sign_off_text = '\n'.join(last_lines).lower()
        
        if '-dan' in sign_off_text:
            return 'dash_dan'
        elif 'thanks, -dan' in sign_off_text:
            return 'thanks_dash_dan'
        elif 'looking forward' in sign_off_text and '-dan' in sign_off_text:
            return 'looking_forward_dan'
        elif 'hope that helps' in sign_off_text and '-dan' in sign_off_text:
            return 'hope_helps_dan'
        else:
            return 'other'

    def count_personal_references(self, content: str) -> int:
        """Count personal context references."""
        personal_patterns = [
            r'\bfamily\b', r'\btravel\b', r'\bmoving\b', r'\bmx\b', r'\bmexico\b',
            r'\bseattle\b', r'\bsummer\b', r'\bbusy\b', r'\bkids\b', r'\bhome\b'
        ]
        
        count = 0
        for pattern in personal_patterns:
            count += len(re.findall(pattern, content, re.IGNORECASE))
        return count

    def count_helping_phrases(self, content: str) -> int:
        """Count helping/service-oriented phrases."""
        helping_patterns = [
            r"anything i can help",
            r"i'd love to help",
            r"would love to help",
            r"happy to help",
            r"let me know if",
            r"is there.*i can help"
        ]
        
        count = 0
        for pattern in helping_patterns:
            count += len(re.findall(pattern, content, re.IGNORECASE))
        return count

    def count_collaborative_language(self, content: str) -> int:
        """Count collaborative language patterns."""
        collaborative_patterns = [
            r"let's\s+\w+",
            r"we could",
            r"what if we",
            r"would love to connect",
            r"i'd love to",
            r"what do you think"
        ]
        
        count = 0
        for pattern in collaborative_patterns:
            count += len(re.findall(pattern, content, re.IGNORECASE))
        return count

    def calculate_avg_sentence_length(self, content: str) -> float:
        """Calculate average sentence length in words."""
        sentences = re.split(r'[.!?]+', content)
        sentences = [s.strip() for s in sentences if s.strip()]
        
        if not sentences:
            return 0.0
            
        total_words = sum(len(sentence.split()) for sentence in sentences)
        return total_words / len(sentences)

    def process_email_batch(self, email_batch: List[Dict]) -> None:
        """Process a batch of emails and add to dataset."""
        existing_contents = [email['raw_data'].get('content', '') for email in self.filtered_emails]
        
        for email_data in email_batch:
            self.stats['total_emails'] += 1
            
            # Skip if repetitive/templated
            if self.is_repetitive_email(email_data.get('subject', ''), 
                                      email_data.get('content', '')):
                self.stats['filtered_out'] += 1
                continue
                
            # Skip if too similar to existing content
            if self.is_duplicate_content(email_data.get('content', ''), existing_contents):
                self.stats['filtered_out'] += 1
                continue
                
            # Process conversational email
            processed_email = {
                'raw_data': email_data,
                'metadata': self.extract_email_metadata(email_data),
                'category': self.categorize_email(email_data),
                'voice_patterns': self.analyze_voice_patterns(email_data),
                'processed_date': datetime.now().isoformat()
            }
            
            self.filtered_emails.append(processed_email)
            existing_contents.append(email_data.get('content', ''))
            self.stats['conversational_emails'] += 1
            
            # Update recipient stats
            to_field = email_data.get('to', '')
            if to_field:
                self.stats['recipients'][to_field] += 1

    def save_dataset(self, output_path: str) -> None:
        """Save the processed dataset to JSON file."""
        dataset = {
            'metadata': {
                'created_date': datetime.now().isoformat(),
                'extraction_period': '6_months',
                'total_emails_processed': self.stats['total_emails'],
                'conversational_emails': self.stats['conversational_emails'],
                'filtered_out': self.stats['filtered_out'],
                'filter_rate': self.stats['filtered_out'] / self.stats['total_emails'] if self.stats['total_emails'] > 0 else 0
            },
            'statistics': self.stats,
            'emails': self.filtered_emails
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(dataset, f, indent=2, ensure_ascii=False)
        
        print(f"Dataset saved to: {output_path}")
        print(f"Total emails processed: {self.stats['total_emails']}")
        print(f"Conversational emails: {self.stats['conversational_emails']}")
        print(f"Filtered out: {self.stats['filtered_out']}")
        print(f"Filter rate: {self.stats['filtered_out'] / self.stats['total_emails'] * 100:.1f}%")

def main():
    """Main execution function - will be called from separate script."""
    extractor = EmailVoiceExtractor()
    
    print("Email Voice Extractor initialized.")
    print("Note: This script defines the processing logic.")
    print("Run 'email_extraction_runner.py' to execute the actual extraction.")
    
    return extractor

if __name__ == "__main__":
    main()
