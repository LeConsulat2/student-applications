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

# --- 설정 ---
load_dotenv()
USERNAME = os.getenv("LNA_USERNAME")
PASSWORD = os.getenv("LNA_PASSWORD")

# 스캔할 페이지 수 (필요에 따라 조절)
MAX_PAGES_TO_SCAN = 11 

if not USERNAME or not PASSWORD:
    print("Error: .env 파일에서 아이디/비번을 찾을 수 없습니다.")
    exit()

def gentle_scroll_to_element(driver, element):
    """요소가 화면 중앙에 오도록 부드럽게 스크롤"""
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
    time.sleep(1)

# --- 메인 로직 시작 ---
print("브라우저를 실행합니다...")

options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_argument("--disable-blink-features=AutomationControlled") 
# detach 옵션을 제거했습니다 (이제 끝나면 닫힙니다)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
wait = WebDriverWait(driver, 20)

try:
    # ==========================================
    # 1단계: 로그인 & 역할 선택
    # ==========================================
    print("--- 1단계: 로그인 ---")
    driver.get("https://assess.literacyandnumeracyforadults.com/Login.aspx")

    # 로그인 버튼 클릭
    wait.until(EC.element_to_be_clickable((By.ID, "ctl00_cphLoginContent_lnkEsaaLogin"))).click()
    time.sleep(3) 

    # 영어/마오리어 체크
    try:
        english_btn = driver.find_elements(By.XPATH, "//button[contains(text(), 'View in English')] | //a[contains(text(), 'View in English')]")
        if english_btn:
            english_btn[0].click()
            time.sleep(2)
    except Exception:
        pass

    # 아이디/비번 입력
    print("아이디/비번 입력 중...")
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

    # 역할 선택 (나올 경우에만)
    print("역할 선택 확인 중...")
    try:
        org_admin_radio = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "ctl00_cphLeftPane_rblRoleChoices_1"))
        )
        print("Organisation Administrator 선택...")
        org_admin_radio.click()
        time.sleep(0.5)
        driver.find_element(By.ID, "ctl00_cphLeftPane_btnSubmit").click()
        print("역할 선택 완료.")
        time.sleep(3) 
    except Exception:
        print("역할 선택 페이지가 없으므로 넘어갑니다.")

    # ==========================================
    # 2단계: Assessments 탭 이동 및 체크
    # ==========================================
    print("\n--- 2단계: 목록 체크 시작 ---")
    
    print("Assessments 탭 클릭...")
    try:
        assessments_tab = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_btnAssessments_imgImage")))
        assessments_tab.click()
    except:
        driver.find_element(By.PARTIAL_LINK_TEXT, "Assessments").click()

    wait.until(EC.presence_of_element_located((By.ID, "ctl00_ctl00_cphMainContent_cphLeftPane_ucAssessments_viewAll_grdAssessments")))
    print("리스트 로딩 완료.")

    current_page = 1
    
    while current_page <= MAX_PAGES_TO_SCAN:
        print(f"\n현재 페이지: {current_page} 처리 중...")
        time.sleep(3) 

        # 체크박스 찾기 및 클릭
        rows = driver.find_elements(By.XPATH, "//tr[.//input[contains(@id, 'SelectCheckBox')]]")
        checked_count = 0
        
        for row in rows:
            try:
                text = row.text.upper()
                
                # 조건: "READ" 또는 "MATH"가 있고 + "(CYCLE" 문구가 있어야 함
                if ("READ" in text or "MATH" in text) and "(CYCLE" in text:
                    checkbox = row.find_element(By.XPATH, ".//input[contains(@id, 'SelectCheckBox')]")
                    if not checkbox.is_selected():
                        checkbox.click()
                        checked_count += 1
                        time.sleep(random.uniform(0.1, 0.3))
            except Exception:
                continue

        print(f"  -> {checked_count}개 선택함.")

        # 페이지 넘기기 로직
        if current_page >= MAX_PAGES_TO_SCAN:
            break

        next_page_num = current_page + 1
        
        try:
            # 숫자(예: 2, 3)가 <span> 안에 들어있는 링크 찾기
            xpath_query = f"//a[.//span[text()='{next_page_num}']]"
            next_page_link = driver.find_element(By.XPATH, xpath_query)
            
            # 스크롤해서 보이게 하고 강제 클릭
            gentle_scroll_to_element(driver, next_page_link)
            driver.execute_script("arguments[0].click();", next_page_link)
            
            print(f"{next_page_num}페이지로 넘어갑니다...")
            time.sleep(5) # 페이지 로딩 대기
            current_page += 1
            
        except Exception:
            print(f"더 이상 {next_page_num}페이지가 없습니다. 선택 종료.")
            break

    # ==========================================
    # 3단계: 엑셀 추출 및 다운로드
    # ==========================================
    print("\n--- 3단계: 다운로드 ---")
    
    # 맨 아래로 스크롤
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
    time.sleep(1)
    
    print("'Get Extract' 버튼 클릭...")
    extract_btn = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphMainContent_cphRightPane_DataEngineReports_lnkDemographicsReport_lnkBtn")))
    extract_btn.click()

    print("다운로드 링크 생성 대기 중...")
    try:
        download_link = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_cphMainContent_cphLeftPane_lnkDownload")))
        print("링크 발견! 클릭합니다.")
        time.sleep(1)
        download_link.click()
        print("성공: 다운로드가 시작되었습니다.")
        
        # --- 안전하게 종료하기 ---
        print("파일 다운로드를 위해 15초간 대기합니다...")
        time.sleep(15) 
        
    except Exception as e:
        print("다운로드 링크를 찾지 못했습니다 (시간 초과).")
        print(e)

except Exception as e:
    print(f"오류 발생: {e}")

finally:
    # 작업이 끝나거나 에러가 나도 브라우저를 닫습니다.
    print("브라우저를 종료합니다.")
    driver.quit()