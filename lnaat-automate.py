import os
import time
import random  # <--- Added this to generate random delays
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
    wait = WebDriverWait(driver, 15) # Increased wait time slightly

    try:
        # --- LOGIN ---
        print("Navigating to login page...")
        driver.get("https://assess.literacyandnumeracyforadults.com/Login.aspx")

        print("Logging in...")
        # Generic locator for username (often 'ctl00_MainContent_txtUsername' but varies)
        user_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[contains(@name, 'sername') or contains(@id, 'sername')]")))
        user_field.clear()
        user_field.send_keys(USERNAME)
        
        # Human-like pause
        time.sleep(random.uniform(0.5, 1.5))

        pass_field = driver.find_element(By.XPATH, "//input[@type='password']")
        pass_field.clear()
        pass_field.send_keys(PASSWORD)

        # Click Login
        login_btn = driver.find_element(By.XPATH, "//input[@type='submit'] | //button[contains(text(), 'Log')]")
        login_btn.click()

        # --- NAVIGATE TO ASSESSMENTS ---
        print("Login successful. Going to Assessments...")
        wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Assessments")))
        driver.get("https://assess.literacyandnumeracyforadults.com/ViewAssessments.aspx")

        # --- SELECTION LOOP ---
        while True:
            # Wait for grid to load
            wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
            time.sleep(2) # Let the page "settle"

            # Find rows
            rows = driver.find_elements(By.XPATH, "//tr[.//input[@type='checkbox']]")
            print(f"Scanning page...")

            count_checked = 0
            for row in rows:
                try:
                    text = row.text
                    if "(Cycle" in text:
                        checkbox = row.find_element(By.XPATH, ".//input[@type='checkbox']")
                        if not checkbox.is_selected():
                            checkbox.click()
                            count_checked += 1
                            
                            # === SAFETY FEATURE ===
                            # Sleep for 0.1 to 0.4 seconds between clicks.
                            # This prevents sending 50 requests in 1 second.
                            time.sleep(random.uniform(0.1, 0.4))
                except Exception:
                    continue 

            print(f"Ticked {count_checked} items on this page.")

            # --- PAGINATION ---
            try:
                # Find the 'Next' arrow (>)
                next_page_btn = driver.find_elements(By.XPATH, "//a[text()='>']")
                
                if next_page_btn and next_page_btn[0].get_attribute('href'):
                    print("Moving to next page...")
                    next_page_btn[0].click()
                    
                    # WAIT longer for page loads (4-6 seconds) to be safe
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

        # --- DOWNLOAD ---
        print("Waiting for download link...")
        download_link = wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Click here to download")))
        
        # Small pause before final click
        time.sleep(2)
        download_link.click()
        
        print("SUCCESS: Download started.")
        time.sleep(10) # Keep browser open to finish download

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()