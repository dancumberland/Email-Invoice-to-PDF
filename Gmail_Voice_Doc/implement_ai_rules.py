#!/usr/bin/env python3
"""
Inbox Zero AI Rules Implementation Script
Automatically sets up advanced AI email management rules in your Inbox Zero system
"""

import requests
import json
import yaml
import time
from typing import Dict, List, Any

class InboxZeroRulesManager:
    def __init__(self, base_url: str = "http://localhost:3000"):
        self.base_url = base_url
        self.session = requests.Session()
        
    def load_rules_config(self, config_path: str) -> Dict[str, Any]:
        """Load the AI rules configuration from YAML file"""
        with open(config_path, 'r') as file:
            return yaml.safe_load(file)
    
    def create_ai_rule(self, rule_config: Dict[str, Any]) -> bool:
        """Create an AI rule in Inbox Zero via API"""
        try:
            # Convert our YAML rule format to Inbox Zero API format
            inbox_zero_rule = self.convert_to_inbox_zero_format(rule_config)
            
            # Make API call to create rule
            response = self.session.post(
                f"{self.base_url}/api/ai/rule",
                json=inbox_zero_rule,
                headers={"Content-Type": "application/json"}
            )
            
            if response.status_code == 200:
                print(f"✅ Successfully created rule: {rule_config['name']}")
                return True
            else:
                print(f"❌ Failed to create rule: {rule_config['name']} - {response.status_code}")
                return False
                
        except Exception as e:
            print(f"❌ Error creating rule {rule_config['name']}: {str(e)}")
            return False
    
    def convert_to_inbox_zero_format(self, rule_config: Dict[str, Any]) -> Dict[str, Any]:
        """Convert our YAML rule format to Inbox Zero's expected format"""
        
        # Build the rule instructions for Claude
        instructions = self.build_rule_instructions(rule_config)
        
        # Create the Inbox Zero rule format
        inbox_zero_rule = {
            "name": rule_config["name"],
            "instructions": instructions,
            "actions": self.convert_actions(rule_config.get("actions", [])),
            "automate": True,
            "runOnThreads": False
        }
        
        return inbox_zero_rule
    
    def build_rule_instructions(self, rule_config: Dict[str, Any]) -> str:
        """Build natural language instructions for Claude"""
        name = rule_config["name"]
        description = rule_config.get("description", "")
        conditions = rule_config.get("conditions", [])
        actions = rule_config.get("actions", [])
        
        instructions = f"""
Rule: {name}
Description: {description}

When to apply this rule:
"""
        
        # Add conditions
        for condition in conditions:
            if condition["type"] == "sender_category":
                instructions += f"- Email is from a {condition['value']}\n"
            elif condition["type"] == "days_since_last_reply":
                instructions += f"- No reply from me in {condition['value']}+ days\n"
            elif condition["type"] == "content_contains":
                keywords = ", ".join(condition["keywords"])
                instructions += f"- Email content contains: {keywords}\n"
            elif condition["type"] == "urgency_indicators":
                keywords = ", ".join(condition["keywords"])
                instructions += f"- Email contains urgency indicators: {keywords}\n"
            elif condition["type"] == "not_labeled":
                instructions += f"- Email is not already labeled with {condition['value']}\n"
        
        instructions += "\nActions to take:\n"
        
        # Add actions
        for action in actions:
            if action["type"] == "add_label":
                instructions += f"- Add Gmail label: {action['value']}\n"
            elif action["type"] == "draft_response":
                template = action.get("template", "professional")
                voice_style = action.get("voice_style", "professional")
                instructions += f"- Draft a {voice_style} response using the {template} template\n"
            elif action["type"] == "set_reminder":
                days = action.get("days", 3)
                instructions += f"- Set a reminder to follow up in {days} days\n"
            elif action["type"] == "extract_context":
                fields = ", ".join(action.get("fields", []))
                instructions += f"- Extract and note: {fields}\n"
            elif action["type"] == "analyze_priority":
                instructions += f"- Analyze email priority based on sender importance and urgency\n"
            elif action["type"] == "flag_if_client":
                instructions += f"- Flag as high priority if from a client domain\n"
        
        instructions += f"""
Always maintain Dan's authentic voice and communication style:
- Professional but approachable
- Direct and clear
- Helpful and solution-oriented
- Personal touch when appropriate
"""
        
        return instructions.strip()
    
    def convert_actions(self, actions: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Convert our action format to Inbox Zero's action format"""
        inbox_zero_actions = []
        
        for action in actions:
            if action["type"] == "add_label":
                inbox_zero_actions.append({
                    "type": "label",
                    "label": action["value"]
                })
            elif action["type"] == "draft_response":
                inbox_zero_actions.append({
                    "type": "reply",
                    "content": f"Draft a response using {action.get('voice_style', 'professional')} tone"
                })
            elif action["type"] == "set_reminder":
                # Inbox Zero might handle this differently, but we'll add as a note
                inbox_zero_actions.append({
                    "type": "forward",
                    "to": "dan.cumberland@gmail.com",
                    "content": f"Reminder: Follow up on this email in {action.get('days', 3)} days"
                })
        
        return inbox_zero_actions
    
    def setup_all_rules(self, config_path: str):
        """Set up all AI rules from the configuration file"""
        print("🚀 Setting up advanced AI rules for Inbox Zero...")
        print("=" * 50)
        
        try:
            config = self.load_rules_config(config_path)
            rules = config.get("ai_rules", [])
            
            success_count = 0
            total_rules = len(rules)
            
            for rule in rules:
                if self.create_ai_rule(rule):
                    success_count += 1
                time.sleep(1)  # Rate limiting
            
            print("\n" + "=" * 50)
            print(f"🎉 Setup Complete!")
            print(f"✅ Successfully created: {success_count}/{total_rules} rules")
            
            if success_count == total_rules:
                print("\n🎯 Your AI email assistant is now ready with advanced automation!")
                print("📧 Rules will automatically process new emails using Claude AI")
                print("🌐 Monitor and adjust rules at: http://localhost:3000")
            else:
                print(f"\n⚠️  {total_rules - success_count} rules failed to create")
                print("💡 You can manually create these rules in the Inbox Zero interface")
                
        except Exception as e:
            print(f"❌ Error setting up rules: {str(e)}")
            print("💡 You can manually create rules using the configuration as a guide")

def main():
    """Main function to run the rules setup"""
    config_path = "/Users/dancumberland/Documents/Work/AI Projects & Training Docs/Dan_Personal_Brand/AI_Tools/Gmail_Voice_Doc/inbox_zero_ai_rules.yaml"
    
    print("🤖 Inbox Zero AI Rules Implementation")
    print("Setting up intelligent email automation...")
    print()
    
    # Check if Inbox Zero is running
    try:
        response = requests.get("http://localhost:3000", timeout=5)
        if response.status_code != 200:
            print("❌ Inbox Zero is not running at http://localhost:3000")
            print("Please start Inbox Zero first, then run this script")
            return
    except requests.exceptions.RequestException:
        print("❌ Cannot connect to Inbox Zero at http://localhost:3000")
        print("Please ensure Inbox Zero is running first")
        return
    
    # Set up the rules
    manager = InboxZeroRulesManager()
    manager.setup_all_rules(config_path)
    
    print("\n📚 Next Steps:")
    print("1. Login to http://localhost:3000")
    print("2. Review and test your AI rules")
    print("3. Adjust rule parameters as needed")
    print("4. Monitor rule performance and effectiveness")
    print("\n🎯 Your intelligent email automation is now active!")

if __name__ == "__main__":
    main()
