import time
import sys
from seleniumbase import Driver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def close_unexpected_ads(driver, allowed_handles):
    """Forcefully terminates any external rogue popup windows/tabs spawned by ads."""
    try:
        current_handles = driver.window_handles
        if len(current_handles) > len(allowed_handles):
            print("⚠️ Rogue external ad tab detected! Closing it...")
            for handle in current_handles:
                if handle not in allowed_handles:
                    try:
                        driver.switch_to.window(handle)
                        driver.close()
                        print("💥 External ad tab terminated.")
                    except Exception:
                        pass
            driver.switch_to.window(allowed_handles[0])
    except Exception:
        pass

def handle_in_page_ads(driver):
    """Looks for and aggressively clicks the specified in-page ad dismissal element."""
    try:
        # Target your specific ad dismissal button xpath
        ad_close_buttons = driver.find_elements(By.XPATH, '//*[@id="dismiss-button-element"]/div')
        for btn in ad_close_buttons:
            if btn.is_displayed():
                print("💥 In-page ad detected! Smacking the dismiss button down...")
                driver.execute_script("arguments[0].click();", btn)
    except Exception:
        pass

def keep_window_alive_for_debugging(driver, status_msg):
    """Prevents the browser session from auto-terminating so you can inspect results."""
    print(f"\n🖥️ {status_msg}")
    print("The browser window has been frozen. Close your terminal to exit.")
    while True:
        try:
            _ = driver.title
            time.sleep(1)
        except Exception:
            print("Browser closed manually. Exiting script.")
            sys.exit()

def main():
    print("🚀 Launching stealth browser via SeleniumBase...")
    driver = Driver(uc=True)
    
    print("📺 Maximizing browser window to ensure coordinate accuracy...")
    driver.maximize_window()
    
    wait = WebDriverWait(driver, 15)
    tab1_handle = driver.current_window_handle
    allowed_tabs = [tab1_handle]

    # ==========================================
    # STEP 1: LOAD TEMP-MAIL & COPY FIRST
    # ==========================================
    print("📅 Navigating to TempMailo...")
    driver.get("https://tempmailo.com/")
    
    copied_text = ""
    try:
        print("⏳ Waiting for email to generate...")
        email_input = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="i-email"]')))
        time.sleep(2.5) 
        copied_text = email_input.get_attribute("value")
        print(f"✅ SUCCESS! Copied email: '{copied_text}'")
        
    except Exception as e:
        keep_window_alive_for_debugging(driver, f"CRITICAL ERROR: Failed to copy the email from Tab 1: {e}")

    # ==========================================
    # STEP 2: CREATE CHATGPT TAB
    # ==========================================
    print("\n🌿 Copy phase successful. Proceeding to open ChatGPT tab...")
    driver.execute_script("window.open('');")
    
    all_tabs = driver.window_handles
    for tab in all_tabs:
        if tab != tab1_handle:
            tab2_handle = tab

    allowed_tabs.append(tab2_handle)
    driver.switch_to.window(tab2_handle)
    print(f"Switched to Tab 2 ID: {tab2_handle}")

    print("🌐 Loading ChatGPT...")
    driver.get("https://www.chatgpt.com")

    # ==========================================
    # STEP 3: DYNAMIC PASTE & PYAUTOGUI CLICK
    # ==========================================
    try:
        import pyautogui
    except ImportError:
        print("❌ Error: pyautogui not found. Run 'pip install pyautogui' first.")
        keep_window_alive_for_debugging(driver, "Missing dependencies.")

    try:
        print("⏳ Hunting for ChatGPT 'Log in' button...")
        login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Log in')]")))
        driver.execute_script("arguments[0].click();", login_button)
        print("🎯 Log In Button successfully clicked!")
        
        print("⏳ Waiting for the authentication form layout to load...")
        time.sleep(3) 
        
        print("🔍 Scanning page for active input boxes...")
        all_inputs = driver.find_elements(By.TAG_NAME, "input")
        username_field = None
        
        for ip in all_inputs:
            if ip.is_displayed() and ip.is_enabled():
                input_type = ip.get_attribute("type")
                if input_type in ["text", "email"] or not input_type:
                    username_field = ip
                    break

        if not username_field:
            print("⚠️ Smart search couldn't verify a visible input. Using structural selector fallback...")
            username_field = wait.until(EC.presence_of_element_located((By.XPATH, "//form//input")))

        try:
            username_field.clear()
        except Exception:
            pass
            
        try:
            username_field.send_keys(copied_text)
            print("📥 Pasted email via standard keys!")
        except Exception:
            driver.execute_script("arguments[0].value = arguments[1];", username_field, copied_text)
            driver.execute_script("arguments[0].dispatchEvent(new Event('input', { bubbles: true }));", username_field)
            print("📥 Pasted email via JavaScript Injection!")
            
        # OS-LEVEL MOUSE SUBMISSION
        print("⏳ Form filled. Triggering OS-level mouse movement...")
        time.sleep(1) 
        
        print("🖱️ Clicking Continue Button at coordinates (667, 680)...")
        pyautogui.moveTo(676, 651, duration=0.5) 
        pyautogui.click()
        print("🚀 Form submitted successfully via hardware simulation!")
        
    except Exception as e:
        keep_window_alive_for_debugging(driver, f"CRITICAL ERROR: Paste/Submit sequence broke: {e}")

    # ==========================================
    # STEP 4: MONITORING LOOP FOR VERIFICATION MAIL
    # ==========================================
    print("\n👀 Transitioning to Active Mail Inbox Monitoring Loop...")
    print("🔄 Shifting context back to Tab 1 (TempMailo)...")
    
    while True:
        try:
            # 1. Clear out rogue pop-up browser tabs instantly
            close_unexpected_ads(driver, allowed_tabs)
            
            # 2. Re-establish focus onto the workspace tab
            driver.switch_to.window(tab1_handle)
            
            # 3. Target and crush any active matching ad structures on the viewport
            handle_in_page_ads(driver)
            
            # 4. Windshield-wiper script to sweep structural ad script containers
            driver.execute_script("""
                var ads = document.querySelectorAll('iframe, [id*="google_ads"], [class*="adsbygoogle"], .ads-wrapper, [id*="pop"]');
                for (var i = 0; i < ads.length; i++) {
                    ads[i].remove();
                }
            """)
            
            # 5. Look for your exact incoming email list container target element
            target_mail_element = driver.find_elements(By.XPATH, '//*[@id="apptmo"]/div/div[2]/div[1]/ul/li/div[1]')
            
            if len(target_mail_element) > 0 and target_mail_element[0].is_displayed():
                print("🚨 SUCCESS: New verification message captured in TempMailo inbox!")
                print("🖱️ Clicking verification email item box...")
                
                # Execute direct JavaScript click to guarantee bypass of transparent ad overlays
                driver.execute_script("arguments[0].click();", target_mail_element[0])
                print("📬 Email opened successfully!")
                break
                
        except Exception:
            # Silently loops through common element stale exceptions during page refreshes
            pass

        time.sleep(2.5)

    # ==========================================
    # STEP 5: FINAL SUCCESS FREEZE 
    # ==========================================
    keep_window_alive_for_debugging(driver, "SUCCESS: Full automation pipeline completed. Mail captures loaded successfully!")

if __name__ == "__main__":
    main()