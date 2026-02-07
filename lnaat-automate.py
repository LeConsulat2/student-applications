import os
import time
import random
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys 
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURATION ---
load_dotenv()
USERNAME = os.getenv("LNA_USERNAME")
PASSWORD = os.getenv("LNA_PASSWORD")

# How many pages do you want to scan?
MAX_PAGES_TO_SCAN = 10 

if not USERNAME or not PASSWORD:
    print("Error: Username or Password not found in .env file.")
    exit()

def gentle_scroll_down(driver):
    """Scrolls down the page in small steps to mimic a human."""
    total_height = driver.execute_script("return document.body.scrollHeight")
    # Scroll in chunks of 300 pixels
    for i in range(0, total_height, 300):
        driver.execute_script(f"window.scrollTo(0, {i});")
        time.sleep(0.1)
    # Ensure we are at the very bottom
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

def main():
    print("Launching browser...")
    
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled") 
    options.add_experimental_option("detach", True) 

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)

    try:
        # ==========================================
        # PART 1: LOGIN & ROLE SELECTION
        # ==========================================
        print("--- PART 1: LOGIN ---")
        driver.get("https://assess.literacyandnumeracyforadults.com/Login.aspx")

        # 1. Click Education Sector Login
        wait.until(EC.element_to_be_clickable((By.ID, "ctl00_cphLoginContent_lnkEsaaLogin"))).click()
        time.sleep(3) 

        # 2. Handle Language (English)
        try:
            english_btn = driver.find_elements(By.XPATH, "//button[contains(text(), 'View in English')] | //a[contains(text(), 'View in English')]")
            if english_btn:
                english_btn[0].click()
                time.sleep(2)
        except Exception:
            pass

        # 3. Enter Credentials
        print("Entering credentials...")
        user_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='text'], input[type='email']")))
        user_field.click()
        user_field.clear()
        for char in USERNAME:
            user_field.send_keys(char)
            time.sleep(0.05) 
        user_field.send_keys(Keys.TAB)

        pass_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        pass_field.click()
        pass_field.clear()
        for char in PASSWORD:
            pass_field.send_keys(char)
            time.sleep(0.05)
        
        time.sleep(1)

        # 4. Click Login (Exact ID)
        print("Clicking Login...")
        wait.until(EC.element_to_be_clickable((By.ID, "loginForm-login"))).click()

        # 5. Handle Role Selection (If it appears)
        print("Checking for Role Selection...")
        try:
            org_admin_radio = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "ctl00_cphLeftPane_rblRoleChoices_1"))
            )
            print("Selecting 'Organisation Administrator'...")
            org_admin_radio.click()
            time.sleep(0.5)
            driver.find_element(By.ID, "ctl00_cphLeftPane_btnSubmit").click()
            print("Role submitted.")
            time.sleep(3) 
        except Exception:
            print("No Role Selection needed. Continuing...")

        # ==========================================
        # PART 2: NAVIGATE TO ASSESSMENTS
        # ==========================================
        print("\n--- PART 2: ASSESSMENTS ---")
        
        # Ensure we are on the dashboard, then click the Assessments Tab Image
        # Using the ID you provided: ctl00_ctl00_btnAssessments_imgImage
        print("Clicking 'Assessments' tab...")
        try:
            assessments_tab = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_btnAssessments_imgImage")))
            assessments_tab.click()
        except:
            # Fallback if ID changes: try link text
            driver.find_element(By.PARTIAL_LINK_TEXT, "Assessments").click()

        # Wait for the grid to load
        wait.until(EC.presence_of_element_located((By.ID, "ctl00_ctl00_cphMainContent_cphLeftPane_ucAssessments_viewAll_grdAssessments")))
        print("Assessment grid loaded.")

        # ==========================================
        # PART 3: THE HARVEST LOOP (Checkboxes)
        # ==========================================
        
        current_page = 1
        
        while current_page <= MAX_PAGES_TO_SCAN:
            print(f"\nProcessing PAGE {current_page}...")
            time.sleep(2) # Allow table to settle

            # 1. Find all rows in the main table
            # We look for TRs that contain a checkbox inside them
            rows = driver.find_elements(By.XPATH, "//tr[.//input[contains(@id, 'SelectCheckBox')]]")
            
            checked_count = 0
            for row in rows:
                try:
                    text = row.text.upper() # Convert to UPPERCASE for easier matching
                    
                    # --- THE FILTERING LOGIC ---
                    # Rule: Must have ("READ" OR "MATH") AND "(CYCLE"
                    # This excludes "Numeracy Session", "Session 8", etc.
                    is_valid_type = "READ" in text or "MATH" in text
                    is_cycle = "(CYCLE" in text
                    
                    if is_valid_type and is_cycle:
                        # Find the checkbox inside THIS row
                        checkbox = row.find_element(By.XPATH, ".//input[contains(@id, 'SelectCheckBox')]")
                        
                        if not checkbox.is_selected():
                            checkbox.click()
                            checked_count += 1
                            # Human delay between ticks
                            time.sleep(random.uniform(0.2, 0.5))
                        else:
                            # Already checked
                            pass
                except Exception:
                    continue # Skip row if stale/error

            print(f"  -> Ticked {checked_count} items on Page {current_page}.")

            # 2. Pagination Logic
            # We look for the link that leads to the NEXT page number.
            # E.g. If we are on page 1, we look for a link with text "2".
            next_page_num = current_page + 1
            
            # Scroll down gently to see pagination and buttons
            gentle_scroll_down(driver)
            
            if current_page >= MAX_PAGES_TO_SCAN:
                print("Max pages reached. Stopping selection.")
                break

            try:
                # Look for a pagination link containing the number 'next_page_num'
                # These are usually inside the footer: <a><span>2</span></a>
                next_page_link = driver.find_element(By.XPATH, f"//tr[@class='PagerStyle']//a[contains(text(), '{next_page_num}')] | //tr[@class='PagerStyle']//a[./span[contains(text(), '{next_page_num}')]]")
                
                print(f"Moving to Page {next_page_num}...")
                next_page_link.click()
                
                # IMPORTANT: Wait for the old rows to disappear (stale) or wait for new page load
                # A simple sleep is often safer on these old ASP.NET sites than complex waits
                time.sleep(4) 
                
                current_page += 1
            except Exception:
                print("No more pages found (or link not clickable). Finished selection.")
                break

        # ==========================================
        # PART 4: EXTRACT & DOWNLOAD
        # ==========================================
        print("\n--- PART 3: EXTRACT ---")
        
        # 1. Scroll to bottom right to find 'Get Extract'
        gentle_scroll_down(driver)
        
        print("Clicking 'Get Extract'...")
        # ID: ctl00_ctl00_cphMainContent_cphRightPane_DataEngineReports_lnkDemographicsReport_lnkBtn
        extract_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphMainContent_cphRightPane_DataEngineReports_lnkDemographicsReport_lnkBtn")))
        extract_btn.click()

        print("Waiting for report generation (this may take a moment)...")
        
        # 2. Wait for the "Click here to download" link
        # ID: ctl00_ctl00_cphMainContent_cphLeftPane_lnkDownload
        try:
            download_link = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphMainContent_cphLeftPane_lnkDownload")))
            
            print("Download link appeared!")
            time.sleep(1)
            download_link.click()
            print("SUCCESS: Download started.")
            
            # Keep browser open long enough for download to finish
            time.sleep(15)
            
        except Exception as e:
            print("Timed out waiting for download link. Did the extract fail?")
            print(e)

    except Exception as e:
        print(f"CRITICAL ERROR: {e}")

if __name__ == "__main__":
    main()