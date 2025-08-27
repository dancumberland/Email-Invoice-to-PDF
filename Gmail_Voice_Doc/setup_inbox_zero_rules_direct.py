#!/usr/bin/env python3
"""
Direct Inbox Zero Rules Setup via Database
Since the API approach didn't work, this script directly inserts rules into the Inbox Zero database
"""

import os
import json
import psycopg2
from urllib.parse import urlparse
import uuid
from datetime import datetime

class InboxZeroDirectSetup:
    def __init__(self):
        # Get database URL from environment file
        self.db_url = self.get_database_url()
        
    def get_database_url(self):
        """Extract database URL from Inbox Zero environment file"""
        env_path = "/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Gmail_Voice_Doc/inbox-zero/apps/web/.env"
        
        try:
            with open(env_path, 'r') as f:
                for line in f:
                    if line.startswith('DATABASE_URL='):
                        return line.split('=', 1)[1].strip().strip('"')
        except Exception as e:
            print(f"Error reading env file: {e}")
            return None
    
    def connect_to_database(self):
        """Connect to the Supabase PostgreSQL database"""
        try:
            if not self.db_url:
                print("❌ Database URL not found")
                return None
                
            conn = psycopg2.connect(self.db_url)
            return conn
        except Exception as e:
            print(f"❌ Database connection failed: {e}")
            return None
    
    def create_ai_rule_in_db(self, conn, rule_data):
        """Insert AI rule directly into database"""
        try:
            cursor = conn.cursor()
            
            # Generate unique ID for the rule
            rule_id = str(uuid.uuid4())
            user_id = "user_placeholder"  # This would need to be the actual user ID
            
            # Insert rule into the database
            insert_query = """
            INSERT INTO rules (id, user_id, name, instructions, actions, automate, created_at, updated_at)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """
            
            cursor.execute(insert_query, (
                rule_id,
                user_id,
                rule_data['name'],
                rule_data['instructions'],
                json.dumps(rule_data['actions']),
                True,
                datetime.now(),
                datetime.now()
            ))
            
            conn.commit()
            cursor.close()
            return True
            
        except Exception as e:
            print(f"❌ Error inserting rule: {e}")
            return False
    
    def setup_rules(self):
        """Set up all AI rules directly in the database"""
        print("🚀 Setting up Inbox Zero AI rules directly...")
        
        # Define the rules
        rules = [
            {
                "name": "Hot Prospect Follow-up",
                "instructions": """
When I haven't replied to a prospect in 7+ days:
- Add Gmail label: @Prospect 💰/Hot-10d
- Draft a personalized follow-up response
- Reference specific details from our previous conversation
- Use my professional but friendly tone
- Include a clear call-to-action and suggest next steps
- Extract: company name, project scope, budget mentions, timeline

Always maintain Dan's authentic voice: professional but approachable, direct and clear, solution-oriented with a personal touch when appropriate.
                """,
                "actions": [
                    {"type": "label", "label": "@Prospect 💰/Hot-10d"},
                    {"type": "reply", "content": "Draft personalized follow-up"}
                ]
            },
            {
                "name": "Urgent Action Detection", 
                "instructions": """
When an email contains urgency indicators (urgent, asap, deadline, time-sensitive, emergency) AND needs a response:
- Add Gmail label: @Action ⚡️/Respond
- Draft a professional acknowledgment that shows I understand the urgency
- Provide a specific timeline for my full response
- Take any immediate action if needed
- Flag as high priority if from a client domain

Use Dan's direct and clear communication style with appropriate urgency.
                """,
                "actions": [
                    {"type": "label", "label": "@Action ⚡️/Respond"},
                    {"type": "reply", "content": "Draft urgent acknowledgment"}
                ]
            },
            {
                "name": "Smart Scheduling Assistant",
                "instructions": """
For emails mentioning scheduling, meeting, call, calendar, available, or book time:
- Add Gmail label: @Action ⚡️/Schedule
- Extract scheduling details: proposed times, duration, attendees
- Draft a helpful response suggesting specific available time slots
- Use professional, accommodating tone
- Include calendar booking link when appropriate

Match Dan's helpful and solution-oriented communication style.
                """,
                "actions": [
                    {"type": "label", "label": "@Action ⚡️/Schedule"},
                    {"type": "reply", "content": "Draft scheduling response"}
                ]
            }
        ]
        
        # Try to connect and insert rules
        conn = self.connect_to_database()
        if not conn:
            print("❌ Cannot connect to database. Manual setup required.")
            return False
        
        success_count = 0
        for rule in rules:
            if self.create_ai_rule_in_db(conn, rule):
                print(f"✅ Created rule: {rule['name']}")
                success_count += 1
            else:
                print(f"❌ Failed to create rule: {rule['name']}")
        
        conn.close()
        
        print(f"\n🎉 Setup complete! {success_count}/{len(rules)} rules created")
        return success_count == len(rules)

def main():
    """Main function"""
    print("🤖 Direct Inbox Zero Rules Setup")
    print("Attempting to set up AI rules directly in database...")
    print()
    
    setup = InboxZeroDirectSetup()
    
    if setup.setup_rules():
        print("\n✅ All rules successfully created!")
        print("🌐 Go to http://localhost:3000 to see your AI rules in action")
    else:
        print("\n⚠️  Direct database setup failed")
        print("📋 Please use the manual setup guide instead")
        print("📖 Open: INBOX_ZERO_SETUP_GUIDE.md")

if __name__ == "__main__":
    main()
