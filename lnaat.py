import os
import time
import random
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
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
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.maximize_window()
    wait = WebDriverWait(driver, 20) 

    try:
        # --- STEP 1: INITIAL PAGE ---
        print("Navigating to initial login page...")
        driver.get("https://assess.literacyandnumeracyforadults.com/Login.aspx")

        print("Clicking 'Log in via ESL'...")
        # UPDATED: This now looks for ANY element containing the text, not just a link
        esl_link = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Click here to log in via ESL')]")))
        esl_link.click()

        # --- STEP 2: ESL LOGIN PAGE ---
        print("Waiting for ESL Login page...")
        
        # Wait for username field
        user_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'username') or contains(@id, 'username') or @type='text']")))
        
        time.sleep(random.uniform(1.0, 2.0))
        
        print("Entering credentials...")
        user_field.clear()
        user_field.send_keys(USERNAME)

        pass_field = driver.find_element(By.XPATH, "//input[@type='password']")
        pass_field.clear()
        pass_field.send_keys(PASSWORD)

        # Click Login
        login_btn = driver.find_element(By.XPATH, "//button[contains(text(), 'Login')] | //input[@type='submit']")
        login_btn.click()

        # --- STEP 3: ROLE SELECTION ---
        print("Waiting for Role Selection...")
        # Wait for "Choose A Role" text
        wait.until(EC.presence_of_element_located((By.XPATH, "//h1[contains(text(), 'Choose A Role')] | //legend[contains(text(), 'Choose A Role')]")))
        
        print("Selecting 'Organisation Administrator'...")
        # Find the radio button for Org Admin
        org_admin_radio = driver.find_element(By.XPATH, "//label[contains(text(), 'Organisation Administrator')]/preceding-sibling::input | //input[@aria-label='Organisation Administrator']")
        
        if not org_admin_radio.is_selected():
            org_admin_radio.click()
        
        time.sleep(1)
        
        # Click Submit
        submit_btn = driver.find_element(By.XPATH, "//input[@value='Submit'] | //button[contains(text(), 'Submit')]")
        submit_btn.click()

        # --- STEP 4: NAVIGATE TO ASSESSMENTS ---
        print("Login complete. Navigating to Assessments...")
        
        wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Assessments")))
        
        assessments_tab = driver.find_element(By.PARTIAL_LINK_TEXT, "Assessments")
        assessments_tab.click()

        # --- STEP 5: SELECTION LOOP ---
        while True:
            # Wait for grid
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(2) 

            rows = driver.find_elements(By.XPATH, "//tr[.//input[@type='checkbox']]")
            print(f"Scanning page...")

            count_checked = 0
            for row in rows:
                try:
                    text = row.text
                    # Checks for "Cycle" in the row text
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
            except Exception as e:
                print(f"Pagination stopped: {e}")
                break

        # --- EXTRACT ---
        print("Clicking 'Get Extract'...")
        extract_link = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Get Extract")))
        extract_link.click()

        print("Waiting for download link...")
        download_link = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Click here to download")))
        
        time.sleep(2)
        download_link.click()
        
        print("SUCCESS: Download started.")
        time.sleep(10) 

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        input("Press Enter to close browser...")
        driver.quit()

if __name__ == "__main__":
    main()