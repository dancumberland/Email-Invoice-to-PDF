#!/usr/bin/env python3
"""
Gmail API Email Extractor for Dan Cumberland
This script uses the Gmail API directly to extract 1000 conversational emails
for comprehensive voice training analysis.
"""

import json
import pickle
import os
import base64
import re
from datetime import datetime, timedelta
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email_voice_extractor import EmailVoiceExtractor

# Gmail API scopes
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

class GmailAPIExtractor:
    def __init__(self):
        self.service = None
        self.extractor = EmailVoiceExtractor()
        self.processed_emails = []
        
    def authenticate(self, credentials_path=None):
        """Authenticate with Gmail API using OAuth2."""
        creds = None
        
        # Default credentials path
        if not credentials_path:
            credentials_path = '/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Gmail_Voice_Doc/gmail_credentials.json'
        
        # Check for existing token
        token_path = '/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Gmail_Voice_Doc/gmail_token.pickle'
        if os.path.exists(token_path):
            with open(token_path, 'rb') as token:
                creds = pickle.load(token)
        
        # If there are no (valid) credentials available, let the user log in
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception as e:
                    print(f"Error refreshing credentials: {e}")
                    creds = None
            
            if not creds:
                if os.path.exists(credentials_path):
                    print(f"Using credentials from: {credentials_path}")
                    flow = InstalledAppFlow.from_client_secrets_file(
                        credentials_path, SCOPES)
                    # Use run_local_server for desktop authentication
                    creds = flow.run_local_server(port=0)
                else:
                    print(f"Error: Gmail OAuth credentials not found at {credentials_path}")
                    print("\nTo set up Gmail API access:")
                    print("1. Go to https://console.cloud.google.com/")
                    print("2. Create a new project or select existing")
                    print("3. Enable Gmail API")
                    print("4. Create OAuth 2.0 credentials (Desktop application)")
                    print(f"5. Download and save as: {credentials_path}")
                    return False
            
            # Save the credentials for the next run
            with open(token_path, 'wb') as token:
                pickle.dump(creds, token)
        
        self.service = build('gmail', 'v1', credentials=creds)
        print("✅ Gmail API authentication successful!")
        return True
    
    def get_email_list(self, query, max_results=1000):
        """Get list of email IDs matching the query."""
        try:
            results = self.service.users().messages().list(
                userId='me', 
                q=query, 
                maxResults=max_results
            ).execute()
            
            messages = results.get('messages', [])
            
            # Get additional pages if needed
            while 'nextPageToken' in results and len(messages) < max_results:
                page_token = results['nextPageToken']
                results = self.service.users().messages().list(
                    userId='me', 
                    q=query, 
                    maxResults=max_results - len(messages),
                    pageToken=page_token
                ).execute()
                messages.extend(results.get('messages', []))
            
            print(f"Found {len(messages)} emails matching query: {query}")
            return messages
            
        except Exception as error:
            print(f'An error occurred: {error}')
            return []
    
    def get_email_content(self, message_id):
        """Get full email content for a message ID."""
        try:
            message = self.service.users().messages().get(
                userId='me', 
                id=message_id,
                format='full'
            ).execute()
            
            # Extract email metadata
            headers = message['payload'].get('headers', [])
            subject = ''
            sender = ''
            recipient = ''
            date = ''
            
            for header in headers:
                name = header['name'].lower()
                if name == 'subject':
                    subject = header['value']
                elif name == 'from':
                    sender = header['value']
                elif name == 'to':
                    recipient = header['value']
                elif name == 'date':
                    date = header['value']
            
            # Extract email body
            body = self.extract_body(message['payload'])
            
            return {
                'id': message_id,
                'subject': subject,
                'from': sender,
                'to': recipient,
                'date': date,
                'content': body
            }
            
        except Exception as error:
            print(f'Error getting email {message_id}: {error}')
            return None
    
    def extract_body(self, payload):
        """Extract email body from payload."""
        body = ''
        
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/plain':
                    if 'data' in part['body']:
                        data = part['body']['data']
                        body = base64.urlsafe_b64decode(data).decode('utf-8')
                        break
                elif part['mimeType'] == 'multipart/alternative':
                    body = self.extract_body(part)
                    if body:
                        break
        else:
            if payload['mimeType'] == 'text/plain':
                if 'data' in payload['body']:
                    data = payload['body']['data']
                    body = base64.urlsafe_b64decode(data).decode('utf-8')
        
        return body
    
    def extract_1000_emails(self):
        """Extract 1000 conversational emails from Gmail."""
        print("🚀 Starting Gmail API extraction for 1000 conversational emails...")
        
        if not self.authenticate():
            return False
        
        # Search queries to get comprehensive email coverage
        search_queries = [
            'from:me after:2025-02-01',  # Last 6 months
            'from:me after:2024-08-01 before:2025-02-01',  # Previous 6 months
            'from:me after:2024-02-01 before:2024-08-01',  # Even earlier if needed
            'from:me after:2023-08-01 before:2024-02-01',  # Go back further if needed
        ]
        
        all_message_ids = []
        
        # Get email IDs from all search queries
        print("\n📧 Collecting email IDs from Gmail...")
        for i, query in enumerate(search_queries, 1):
            print(f"  Query {i}: {query}")
            messages = self.get_email_list(query, max_results=2000)
            all_message_ids.extend([msg['id'] for msg in messages])
            print(f"    Found: {len(messages)} emails")
            
            if len(all_message_ids) >= 5000:  # Get plenty to filter from
                print(f"    Collected enough emails ({len(all_message_ids)}), stopping search")
                break
        
        print(f"\n📊 Total email IDs collected: {len(all_message_ids)}")
        
        # Remove duplicates
        unique_message_ids = list(set(all_message_ids))
        print(f"📊 Unique email IDs: {len(unique_message_ids)}")
        
        # Process emails in batches
        batch_size = 25  # Smaller batches for better progress tracking
        conversational_count = 0
        target_count = 1000
        processed_count = 0
        
        print(f"\n🔄 Processing emails in batches of {batch_size}...")
        print(f"🎯 Target: {target_count} conversational emails")
        
        for i in range(0, len(unique_message_ids), batch_size):
            if conversational_count >= target_count:
                break
                
            batch_ids = unique_message_ids[i:i + batch_size]
            batch_num = i//batch_size + 1
            total_batches = len(unique_message_ids) // batch_size + (1 if len(unique_message_ids) % batch_size else 0)
            
            print(f"\n📦 Batch {batch_num}/{total_batches}: Processing emails {i+1}-{min(i+batch_size, len(unique_message_ids))}")
            
            batch_emails = []
            successful_reads = 0
            
            for j, message_id in enumerate(batch_ids, 1):
                try:
                    email_data = self.get_email_content(message_id)
                    if email_data and email_data['content']:
                        batch_emails.append(email_data)
                        successful_reads += 1
                    
                    # Progress indicator within batch
                    if j % 5 == 0 or j == len(batch_ids):
                        print(f"    📖 Read {j}/{len(batch_ids)} emails in batch")
                        
                except Exception as e:
                    print(f"    ⚠️ Error reading email {message_id}: {str(e)[:50]}...")
                    continue
            
            print(f"    ✅ Successfully read {successful_reads}/{len(batch_ids)} emails")
            
            # Process batch through the extractor
            initial_count = len(self.extractor.filtered_emails)
            self.extractor.process_email_batch(batch_emails)
            new_conversational = len(self.extractor.filtered_emails) - initial_count
            conversational_count += new_conversational
            processed_count += len(batch_emails)
            
            # Progress summary
            progress_pct = (conversational_count / target_count) * 100
            print(f"    📈 Batch result: {new_conversational} conversational out of {len(batch_emails)} emails")
            print(f"    🎯 Progress: {conversational_count}/{target_count} ({progress_pct:.1f}%) conversational emails")
            print(f"    📊 Total processed: {processed_count} emails")
            
            if conversational_count >= target_count:
                print(f"\n🎉 TARGET REACHED! {conversational_count} conversational emails extracted!")
                break
        
        # Save the comprehensive dataset
        output_path = "/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Gmail_Voice_Doc/email_voice_dataset_1000.json"
        self.extractor.save_dataset(output_path)
        
        print(f"\n📊 FINAL EXTRACTION RESULTS:")
        print(f"="*50)
        print(f"✅ Total emails processed: {self.extractor.stats['total_emails']}")
        print(f"🎯 Conversational emails found: {self.extractor.stats['conversational_emails']}")
        print(f"🚫 Filtered out (repetitive): {self.extractor.stats['filtered_out']}")
        if self.extractor.stats['total_emails'] > 0:
            filter_rate = self.extractor.stats['filtered_out'] / self.extractor.stats['total_emails'] * 100
            print(f"📈 Filter rate: {filter_rate:.1f}%")
        print(f"💾 Dataset saved to: {output_path}")
        
        # Success indicator
        if conversational_count >= target_count:
            print(f"\n🏆 SUCCESS: Target of {target_count} conversational emails achieved!")
        else:
            print(f"\n⚠️ PARTIAL SUCCESS: Found {conversational_count} conversational emails (target: {target_count})")
        
        return True

def main():
    """Main execution function."""
    extractor = GmailAPIExtractor()
    
    print("Dan Cumberland Gmail API Email Extractor")
    print("="*50)
    print("This script will:")
    print("1. Authenticate with Gmail API using OAuth2")
    print("2. Search for emails from multiple date ranges")
    print("3. Extract and filter 1000 conversational emails")
    print("4. Process through voice analysis")
    print("5. Save comprehensive dataset")
    print()
    
    success = extractor.extract_1000_emails()
    
    if success:
        print("\n✅ Email extraction completed successfully!")
        print("Ready to enhance the voice training document with comprehensive data.")
    else:
        print("\n❌ Email extraction failed. Please check authentication and try again.")

if __name__ == "__main__":
    main()
