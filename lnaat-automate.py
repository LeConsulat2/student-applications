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
    
    # 1. Setup Chrome Options to prevent random crashes
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-blink-features=AutomationControlled") 
    
    # Keep browser open if script crashes (helps debugging)
    options.add_experimental_option("detach", True)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20) # Increased wait to 20s

    try:
        # --- LOGIN ---
        print("Navigating to login page...")
        driver.get("https://assess.literacyandnumeracyforadults.com/Login.aspx")

        # 1. CLICK THE ESAA BUTTON
        print("Clicking Education Sector Login...")
        esaa_button = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_cphLoginContent_lnkEsaaLogin")))
        esaa_button.click()

        # 2. HANDLE THE EDUCATION SECTOR PAGE
        print("Waiting for Education Sector page...")
        time.sleep(3) # Let the redirect happen safely

        # --- FIX: Check for "View in English" button ---
        # The page might load in Māori. If we see "View in English", we click it.
        try:
            english_btn = driver.find_elements(By.XPATH, "//button[contains(text(), 'View in English')] | //a[contains(text(), 'View in English')]")
            if english_btn:
                print("Detected Māori interface. Switching to English...")
                english_btn[0].click()
                time.sleep(2) # Wait for language reload
        except Exception:
            print("Language switch failed or not needed. Continuing...")

        # 3. ENTER CREDENTIALS
        print("Entering credentials...")
        
        # We use generic input types because IDs might change, but the order (User -> Pass) is stable
        # Find the first text input (Username)
        user_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[type='text'], input[type='email']")))
        user_field.clear()
        user_field.send_keys(USERNAME)

        # Find the password field
        pass_field = driver.find_element(By.CSS_SELECTOR, "input[type='password']")
        pass_field.clear()
        pass_field.send_keys(PASSWORD)

        time.sleep(1)

        # Click the Login Button (Looking for type='submit' or the specific class)
        # This works for both "Login" (English) and "Takiuru" (Māori)
        login_btn = driver.find_element(By.CSS_SELECTOR, "button[type='submit'], input[type='submit']")
        login_btn.click()

        # --- NAVIGATE TO ASSESSMENTS ---
        print("Login clicked. Waiting for return to dashboard...")
        
        # Wait for the URL to return to the main site OR 'Assessments' link to appear
        wait.until(EC.url_contains("assess.literacyandnumeracyforadults.com"))
        wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Assessments")))
        
        print("Login successful. Going to Assessments...")
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