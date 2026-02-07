import os
import time
import random
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys # <--- ESSENTIAL IMPORT
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
    
    # 1. Setup Chrome Options
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled") 
    options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)

    try:
        # --- LOGIN PHASE ---
        print("Navigating to login page...")
        driver.get("https://assess.literacyandnumeracyforadults.com/Login.aspx")

        print("Clicking Education Sector Login...")
        esaa_button = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_cphLoginContent_lnkEsaaLogin")))
        esaa_button.click()

        print("Waiting for redirection...")
        time.sleep(3) 

        # Force English if possible (makes debugging easier)
        try:
            english_btn = driver.find_elements(By.XPATH, "//button[contains(text(), 'View in English')] | //a[contains(text(), 'View in English')]")
            if english_btn:
                print("Switching language to English...")
                english_btn[0].click()
                time.sleep(2)
        except Exception:
            pass

        # --- THE FIX: SLOW TYPING MODE ---
        print("Entering credentials slowly...")

        # 1. Username
        user_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='text'], input[type='email']")))
        user_field.click()
        user_field.clear()
        
        # Type one letter at a time
        for char in USERNAME:
            user_field.send_keys(char)
            time.sleep(0.1) # Small delay between keys
        
        # Press TAB to 'lock in' the value
        user_field.send_keys(Keys.TAB)
        time.sleep(1)

        # 2. Password
        pass_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        pass_field.click()
        pass_field.clear()
        
        for char in PASSWORD:
            pass_field.send_keys(char)
            time.sleep(0.1)

        pass_field.send_keys(Keys.TAB)
        time.sleep(1)

        # 3. Click Login
        print("Clicking Login button...")
        login_btn = driver.find_element(By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")
        login_btn.click()

        # --- VERIFY LOGIN ---
        print("Checking if login worked...")
        try:
            # Wait for the URL to change back to the assessment site
            wait.until(EC.url_contains("assess.literacyandnumeracyforadults.com"))
            print("Redirect successful!")
        except:
            print("Warning: URL didn't change quickly. Checking for failure...")

        # Check if we are actually in the dashboard
        wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Assessments")))
        print("Login CONFIRMED. Going to Assessments...")
        
        driver.get("https://assess.literacyandnumeracyforadults.com/ViewAssessments.aspx")

        # --- SELECTION LOOP (Your original logic) ---
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
        # Browser stays open so you can see what happened

if __name__ == "__main__":
    main()