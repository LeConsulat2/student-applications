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

# 1. Load credentials
load_dotenv()
USERNAME = os.getenv("LNA_USERNAME")
PASSWORD = os.getenv("LNA_PASSWORD")

if not USERNAME or not PASSWORD:
    print("Error: Username or Password not found in .env file.")
    exit()

def main():
    print("Launching browser...")
    
    # Setup Chrome Options
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled") 
    options.add_experimental_option("detach", True) 

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)

    try:
        # --- STEP 1: LOGIN PHASE ---
        print("Navigating to login page...")
        driver.get("https://assess.literacyandnumeracyforadults.com/Login.aspx")

        print("Clicking Education Sector Login...")
        esaa_button = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_cphLoginContent_lnkEsaaLogin")))
        esaa_button.click()

        print("Waiting for redirection...")
        time.sleep(3) 

        # Language Switch (Safety Check)
        try:
            english_btn = driver.find_elements(By.XPATH, "//button[contains(text(), 'View in English')] | //a[contains(text(), 'View in English')]")
            if english_btn:
                english_btn[0].click()
                time.sleep(2)
        except Exception:
            pass

        # --- STEP 2: ENTER CREDENTIALS ---
        print("Entering credentials...")

        # Username
        user_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='text'], input[type='email']")))
        user_field.click()
        user_field.clear()
        for char in USERNAME:
            user_field.send_keys(char)
            time.sleep(0.05) 
        user_field.send_keys(Keys.TAB)

        # Password
        pass_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        pass_field.click()
        pass_field.clear()
        for char in PASSWORD:
            pass_field.send_keys(char)
            time.sleep(0.05)
        
        # --- STEP 3: CLICK LOGIN ---
        print("Clicking Login...")
        login_btn = wait.until(EC.element_to_be_clickable((By.ID, "loginForm-login")))
        login_btn.click()

        # --- STEP 4: HANDLE ROLE SELECTION (NEW!) ---
        print("Checking for 'Choose A Role' page...")
        try:
            # We wait up to 5 seconds to see if the "Organisation Administrator" radio button appears.
            # Using the specific ID you provided: ctl00_cphLeftPane_rblRoleChoices_1
            org_admin_radio = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "ctl00_cphLeftPane_rblRoleChoices_1"))
            )
            print("Role selection found. Selecting 'Organisation Administrator'...")
            org_admin_radio.click()
            time.sleep(0.5)

            # Click the Submit button (using the ID you provided)
            submit_role_btn = driver.find_element(By.ID, "ctl00_cphLeftPane_btnSubmit")
            submit_role_btn.click()
            print("Role submitted. Waiting for dashboard...")
            time.sleep(3) # Wait for the redirect to finish
            
        except Exception:
            print("No Role Selection page appeared (or timed out). Proceeding directly...")

        # --- STEP 5: VERIFY DASHBOARD ---
        print("Verifying we are in...")
        wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Assessments")))
        print("Login CONFIRMED. Going to Assessments...")
        
        driver.get("https://assess.literacyandnumeracyforadults.com/ViewAssessments.aspx")

        # --- STEP 6: SELECTION LOOP ---
        while True:
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(2) 

            rows = driver.find_elements(By.XPATH, "//tr[.//input[@type='checkbox']]")
            print(f"Scanning page for checkboxes...")

            count_checked = 0
            for row in rows:
                try:
                    text = row.text
                    if "(Cycle" in text:
                        checkbox = row.find_element(By.XPATH, ".//input[@type='checkbox']")
                        if not checkbox.is_selected():
                            checkbox.click()
                            count_checked += 1
                            time.sleep(random.uniform(0.1, 0.3))
                except Exception:
                    continue 

            print(f"Ticked {count_checked} items.")

            # Pagination
            try:
                next_page_btn = driver.find_elements(By.XPATH, "//a[text()='>']")
                if next_page_btn and next_page_btn[0].get_attribute('href'):
                    print("Moving to next page...")
                    next_page_btn[0].click()
                    time.sleep(random.uniform(4.0, 6.0)) 
                else:
                    print("No 'Next' button found. Finished selection.")
                    break
            except Exception:
                break

        # --- EXTRACT & DOWNLOAD ---
        print("Clicking 'Get Extract'...")
        extract_link = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Get Extract")))
        extract_link.click()

        print("Waiting for download link...")
        download_link = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Click here to download")))
        time.sleep(2)
        download_link.click()
        
        print("SUCCESS: Download started.")
        time.sleep(15) 

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()