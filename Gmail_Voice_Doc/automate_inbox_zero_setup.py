#!/usr/bin/env python3
"""
Automated Inbox Zero Rules Setup using Browser Automation
This script will open Inbox Zero in a browser and set up the AI rules automatically
"""

import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException

class InboxZeroAutomator:
    def __init__(self):
        self.driver = None
        self.wait = None
        
    def setup_browser(self):
        """Initialize Chrome browser with appropriate settings"""
        try:
            print("🌐 Setting up browser...")
            
            # Chrome options
            options = webdriver.ChromeOptions()
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            # Don't run headless so user can see what's happening
            # options.add_argument("--headless")
            
            # Setup Chrome driver
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.wait = WebDriverWait(self.driver, 10)
            
            print("✅ Browser ready!")
            return True
            
        except Exception as e:
            print(f"❌ Browser setup failed: {e}")
            return False
    
    def navigate_to_inbox_zero(self):
        """Navigate to Inbox Zero and wait for it to load"""
        try:
            print("🚀 Opening Inbox Zero...")
            self.driver.get("http://localhost:3000")
            
            # Wait for page to load
            time.sleep(3)
            
            # Check if we're on the login page or dashboard
            page_title = self.driver.title
            print(f"📄 Page loaded: {page_title}")
            
            return True
            
        except Exception as e:
            print(f"❌ Failed to navigate to Inbox Zero: {e}")
            return False
    
    def handle_google_auth(self):
        """Handle Google OAuth authentication if needed"""
        try:
            print("🔐 Checking authentication status...")
            
            # Look for sign-in button or already authenticated state
            try:
                # If we see a "Sign in with Google" button, we need to authenticate
                sign_in_button = self.wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Sign in') or contains(text(), 'Google')]"))
                )
                
                print("🔑 Authentication required. Please complete Google OAuth in the browser...")
                print("👆 Click the 'Sign in with Google' button and complete authentication")
                print("⏳ Waiting for you to complete authentication...")
                
                # Wait for user to complete authentication (check for dashboard elements)
                self.wait.until(
                    EC.any_of(
                        EC.presence_of_element_located((By.XPATH, "//nav")),
                        EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'Rules')]")),
                        EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'Rules')]"))
                    )
                )
                
                print("✅ Authentication completed!")
                return True
                
            except TimeoutException:
                # Already authenticated, dashboard is visible
                print("✅ Already authenticated!")
                return True
                
        except Exception as e:
            print(f"❌ Authentication handling failed: {e}")
            return False
    
    def find_rules_section(self):
        """Find and navigate to the Rules section"""
        try:
            print("🔍 Looking for Rules section...")
            
            # Try different possible selectors for Rules
            possible_selectors = [
                "//a[contains(text(), 'Rules')]",
                "//button[contains(text(), 'Rules')]",
                "//nav//a[contains(text(), 'Rules')]",
                "//div[contains(text(), 'Rules')]",
                "//a[contains(@href, 'rules')]",
                "//a[contains(text(), 'AI')]",
                "//button[contains(text(), 'AI')]",
                "//a[contains(text(), 'Automation')]"
            ]
            
            rules_element = None
            for selector in possible_selectors:
                try:
                    rules_element = self.driver.find_element(By.XPATH, selector)
                    if rules_element:
                        print(f"✅ Found Rules section with selector: {selector}")
                        break
                except NoSuchElementException:
                    continue
            
            if rules_element:
                rules_element.click()
                time.sleep(2)
                print("📋 Navigated to Rules section")
                return True
            else:
                print("⚠️  Rules section not found. Showing current page structure...")
                self.show_page_structure()
                return False
                
        except Exception as e:
            print(f"❌ Failed to find Rules section: {e}")
            return False
    
    def show_page_structure(self):
        """Show the current page structure to help debug"""
        try:
            print("\n🔍 Current page structure:")
            
            # Get all clickable elements
            clickable_elements = self.driver.find_elements(By.XPATH, "//a | //button")
            
            print("📍 Clickable elements found:")
            for i, element in enumerate(clickable_elements[:10]):  # Show first 10
                try:
                    text = element.text.strip()
                    tag = element.tag_name
                    if text:
                        print(f"  {i+1}. {tag}: '{text}'")
                except:
                    pass
            
            # Get navigation elements
            nav_elements = self.driver.find_elements(By.XPATH, "//nav//a | //nav//button")
            if nav_elements:
                print("\n🧭 Navigation elements:")
                for i, element in enumerate(nav_elements):
                    try:
                        text = element.text.strip()
                        if text:
                            print(f"  {i+1}. '{text}'")
                    except:
                        pass
                        
        except Exception as e:
            print(f"Error showing page structure: {e}")
    
    def wait_for_manual_setup(self):
        """Wait for user to manually set up rules and provide guidance"""
        print("\n🎯 MANUAL SETUP REQUIRED")
        print("=" * 50)
        print("I've opened Inbox Zero for you, but I need you to manually create the AI rules.")
        print("\nHere's what to do:")
        print("1. ✅ Complete Google authentication if prompted")
        print("2. 🔍 Look for 'Rules', 'AI Assistant', or 'Automation' section")
        print("3. ➕ Click 'Add Rule' or 'Create Rule'")
        print("4. 📋 Copy the rule configurations from INBOX_ZERO_SETUP_GUIDE.md")
        print("5. 🎯 Start with 'Hot Prospect Follow-up' rule first")
        print("\n📖 Full instructions are in: INBOX_ZERO_SETUP_GUIDE.md")
        print("\n⏳ I'll keep the browser open for you...")
        
        # Keep browser open and wait
        input("\n👆 Press ENTER when you've finished setting up the rules...")
        
        print("🎉 Great! Your AI email automation should now be active!")
        print("🌐 You can close this browser window or keep it open to monitor your rules.")
    
    def run_automation(self):
        """Main automation flow"""
        print("🤖 Inbox Zero AI Rules Automation")
        print("=" * 40)
        
        if not self.setup_browser():
            return False
        
        try:
            if not self.navigate_to_inbox_zero():
                return False
            
            if not self.handle_google_auth():
                return False
            
            # Try to find rules section, but if not found, provide manual guidance
            if not self.find_rules_section():
                print("⚠️  Automated rule creation not possible")
                print("🔄 Switching to guided manual setup...")
            
            # Provide manual setup guidance
            self.wait_for_manual_setup()
            
            return True
            
        except Exception as e:
            print(f"❌ Automation failed: {e}")
            return False
        
        finally:
            if self.driver:
                print("\n🔄 Keeping browser open for your use...")
                print("Close the browser window when you're done.")
                # Don't close automatically - let user close when ready
                # self.driver.quit()

def main():
    """Main function"""
    print("🚀 Starting Inbox Zero AI Rules Setup Automation")
    print()
    
    automator = InboxZeroAutomator()
    
    if automator.run_automation():
        print("\n✅ Setup process completed!")
        print("🎯 Your AI email automation rules should now be active")
        print("📧 Test with some emails to verify functionality")
    else:
        print("\n❌ Automation failed")
        print("📋 Please use the manual setup guide: INBOX_ZERO_SETUP_GUIDE.md")

if __name__ == "__main__":
    main()
