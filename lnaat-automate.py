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

# Set this to 9 or 10 based on your screenshot
MAX_PAGES_TO_SCAN = 10 

if not USERNAME or not PASSWORD:
    print("Error: Username or Password not found in .env file.")
    exit()

def gentle_scroll_to_element(driver, element):
    """Scrolls the element into view so we can see it."""
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
    time.sleep(1)

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
        # PART 1: LOGIN & ROLE
        # ==========================================
        print("--- PART 1: LOGIN ---")
        driver.get("https://assess.literacyandnumeracyforadults.com/Login.aspx")

        # 1. Login Button
        wait.until(EC.element_to_be_clickable((By.ID, "ctl00_cphLoginContent_lnkEsaaLogin"))).click()
        time.sleep(3) 

        # 2. Language Check
        try:
            english_btn = driver.find_elements(By.XPATH, "//button[contains(text(), 'View in English')] | //a[contains(text(), 'View in English')]")
            if english_btn:
                english_btn[0].click()
                time.sleep(2)
        except Exception:
            pass

        # 3. Credentials
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
        wait.until(EC.element_to_be_clickable((By.ID, "loginForm-login"))).click()

        # 4. Role Selection
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
        # PART 2: ASSESSMENTS LOOP
        # ==========================================
        print("\n--- PART 2: ASSESSMENTS ---")
        
        print("Clicking 'Assessments' tab...")
        try:
            assessments_tab = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_btnAssessments_imgImage")))
            assessments_tab.click()
        except:
            driver.find_element(By.PARTIAL_LINK_TEXT, "Assessments").click()

        wait.until(EC.presence_of_element_located((By.ID, "ctl00_ctl00_cphMainContent_cphLeftPane_ucAssessments_viewAll_grdAssessments")))
        print("Assessment grid loaded.")

        current_page = 1
        
        while current_page <= MAX_PAGES_TO_SCAN:
            print(f"\nProcessing PAGE {current_page}...")
            # Wait for grid to stabilize
            time.sleep(3) 

            # --- TICK CHECKBOXES ---
            rows = driver.find_elements(By.XPATH, "//tr[.//input[contains(@id, 'SelectCheckBox')]]")
            checked_count = 0
            
            for row in rows:
                try:
                    text = row.text.upper()
                    
                    # FILTER: Match "READ" or "MATH" AND "(CYCLE"
                    if ("READ" in text or "MATH" in text) and "(CYCLE" in text:
                        checkbox = row.find_element(By.XPATH, ".//input[contains(@id, 'SelectCheckBox')]")
                        if not checkbox.is_selected():
                            checkbox.click()
                            checked_count += 1
                            time.sleep(random.uniform(0.1, 0.3))
                except Exception:
                    continue

            print(f"  -> Ticked {checked_count} items on Page {current_page}.")

            # --- PAGINATION LOGIC (The Fix) ---
            if current_page >= MAX_PAGES_TO_SCAN:
                break

            next_page_num = current_page + 1
            print(f"Looking for button: '{next_page_num}'...")

            try:
                # EXACT MATCH: Look for an 'a' tag that contains a 'span' with the number '2', '3', etc.
                # XPath: //a[.//span[text()='2']]
                xpath_query = f"//a[.//span[text()='{next_page_num}']]"
                
                next_page_link = driver.find_element(By.XPATH, xpath_query)
                
                # Scroll to it just in case
                gentle_scroll_to_element(driver, next_page_link)
                
                print(f"Clicking Page {next_page_num}...")
                
                # FORCE CLICK: Standard .click() fails on these ASP.NET links sometimes
                driver.execute_script("arguments[0].click();", next_page_link)
                
                # Wait for reload (Legacy sites are slow)
                print("Waiting for page reload...")
                time.sleep(5) 
                
                current_page += 1
                
            except Exception as e:
                print(f"Could not find Page {next_page_num}. Assuming end of list.")
                # Optional: Try looking for a "Next >" arrow just in case
                break

        # ==========================================
        # PART 3: EXTRACT
        # ==========================================
        print("\n--- PART 3: EXTRACT ---")
        
        # Scroll down to make sure the extract button is visible
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        
        print("Clicking 'Get Extract'...")
        extract_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphMainContent_cphRightPane_DataEngineReports_lnkDemographicsReport_lnkBtn")))
        extract_btn.click()

        print("Waiting for download link...")
        try:
            download_link = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphMainContent_cphLeftPane_lnkDownload")))
            print("Download link appeared!")
            time.sleep(1)
            download_link.click()
            print("SUCCESS: Download started.")
            time.sleep(15)
        except Exception as e:
            print("Timed out waiting for download link.")
            print(e)

    except Exception as e:
        print(f"CRITICAL ERROR: {e}")

if __name__ == "__main__":
    main()