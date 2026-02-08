import os
import datetime
import pandas as pd
import glob
import warnings
from win32com.mapi import mapi, mapiutil
from win32com.mapi.mapitags import *
import pythoncom
import winreg

# ================= CONFIGURATION =================
user_docs = os.path.join(os.environ["USERPROFILE"], "Documents")
OUTPUT_FOLDER = os.path.join(user_docs, "student-applications")
RAW_FILES_FOLDER = os.path.join(OUTPUT_FOLDER, "files")
MASTER_FILENAME = "Active Application Details.xlsx"
TARGET_SUBJECT = "NSN Report"

START_DATE = datetime.datetime(2025, 9, 1)
END_DATE = datetime.datetime(2026, 2, 5, 23, 59, 59)

warnings.filterwarnings("ignore")
# =================================================

def get_custom_filename(folder: str, date_obj: datetime.datetime) -> str:
    """Creates: YYYY-MM-DD__Active-Applications-001__(DD-MM-YYYY).xlsx"""
    sort_prefix = date_obj.strftime("%Y-%m-%d")
    nz_suffix = date_obj.strftime("%d-%m-%Y")
    counter = 1
    while True:
        filename = f"{sort_prefix}__Active-Applications-{counter:03d}__({nz_suffix}).xlsx"
        full_path = os.path.join(folder, filename)
        if not os.path.exists(full_path):
            return filename
        counter += 1

def get_mapi_profiles():
    """레지스트리에서 MAPI 프로필 목록 가져오기"""
    profiles = []
    try:
        # MAPI 프로필은 레지스트리에 저장됨
        reg_path = r"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
        
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path)
        
        i = 0
        while True:
            try:
                profile_name = winreg.EnumKey(key, i)
                profiles.append(profile_name)
                i += 1
            except OSError:
                break
        
        winreg.CloseKey(key)
        
    except Exception as e:
        print(f"   ⚠️ 레지스트리 읽기 실패: {e}")
    
    return profiles

def get_default_profile():
    """기본 MAPI 프로필 가져오기"""
    try:
        reg_path = r"Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path)
        
        try:
            default_profile, _ = winreg.QueryValueEx(key, "DefaultProfile")
            winreg.CloseKey(key)
            return default_profile
        except:
            winreg.CloseKey(key)
            return None
            
    except Exception:
        return None

def mapi_login():
    """순수 MAPI로 로그인 - 개선 버전"""
    print("🔗 MAPI 직접 로그인 시도...")
    
    try:
        # MAPI 초기화
        mapi.MAPIInitialize(None)
        print("   ✅ MAPI 초기화 완료")
        
        # 사용 가능한 프로필 확인
        profiles = get_mapi_profiles()
        default_profile = get_default_profile()
        
        print(f"\n📋 발견된 MAPI 프로필:")
        if profiles:
            for p in profiles:
                marker = " ⭐ (기본)" if p == default_profile else ""
                print(f"   - {p}{marker}")
        else:
            print("   ❌ 프로필 없음!")
            print("\n해결 방법:")
            print("   1. Windows 검색 → 'Mail' 입력")
            print("   2. 'Mail (Microsoft Outlook)' 실행")
            print("   3. 'Show Profiles' → 'Add' 클릭")
            print("   4. 이메일 계정 추가")
            return None
        
        # 로그인 시도 - 여러 방법 시도
        session = None
        
        # 방법 1: 기본 프로필로 시도
        if default_profile:
            print(f"\n🔑 기본 프로필로 로그인 시도: '{default_profile}'")
            try:
                flags = mapi.MAPI_EXTENDED | mapi.MAPI_NEW_SESSION
                session = mapi.MAPILogonEx(
                    0,
                    default_profile,  # 프로필 이름 명시
                    None,
                    flags
                )
                print("   ✅ 로그인 성공!")
                return session
            except Exception as e:
                print(f"   ❌ 실패: {e}")
        
        # 방법 2: 첫 번째 프로필로 시도
        if profiles and not session:
            print(f"\n🔑 첫 번째 프로필로 시도: '{profiles[0]}'")
            try:
                flags = mapi.MAPI_EXTENDED | mapi.MAPI_NEW_SESSION
                session = mapi.MAPILogonEx(
                    0,
                    profiles[0],
                    None,
                    flags
                )
                print("   ✅ 로그인 성공!")
                return session
            except Exception as e:
                print(f"   ❌ 실패: {e}")
        
        # 방법 3: MAPI_USE_DEFAULT 플래그 사용
        if not session:
            print(f"\n🔑 기본 플래그로 시도...")
            try:
                flags = mapi.MAPI_EXTENDED | mapi.MAPI_USE_DEFAULT | mapi.MAPI_NEW_SESSION
                session = mapi.MAPILogonEx(
                    0,
                    "",  # 빈 문자열
                    "",
                    flags
                )
                print("   ✅ 로그인 성공!")
                return session
            except Exception as e:
                print(f"   ❌ 실패: {e}")
        
        # 방법 4: SimpleMAPI 사용 (폴백)
        if not session:
            print(f"\n🔑 SimpleMAPI로 시도...")
            try:
                flags = mapi.MAPI_NEW_SESSION | mapi.MAPI_LOGON_UI
                session = mapi.MAPILogon(
                    0,
                    None,
                    None,
                    flags,
                    0
                )
                print("   ✅ SimpleMAPI 로그인 성공!")
                return session
            except Exception as e:
                print(f"   ❌ 실패: {e}")
        
        print("\n❌ 모든 로그인 시도 실패")
        print("\n해결 방법:")
        print("1. Classic Outlook을 한 번 실행해서 프로필 초기화")
        print("2. Control Panel → Mail → Show Profiles → 'Outlook' 프로필 확인")
        print("3. 프로필이 없으면 새로 추가")
        return None
        
    except Exception as e:
        print(f"   ❌ MAPI 초기화 실패: {e}")
        import traceback
        traceback.print_exc()
        return None

def get_prop_value(props, prop_tag):
    """속성에서 값 추출"""
    for prop in props:
        if prop[0] == prop_tag:
            return prop[1]
    return None

def get_inbox(session):
    """Inbox 폴더 가져오기"""
    try:
        # Message Store 테이블 가져오기
        msg_stores_table = session.GetMsgStoresTable(0)
        msg_stores_table.SetColumns((PR_ENTRYID, PR_DISPLAY_NAME, PR_DEFAULT_STORE), 0)
        
        print("\n📬 사용 가능한 Message Stores:")
        
        # Default store 찾기
        while True:
            rows = msg_stores_table.QueryRows(1, 0)
            if len(rows) == 0:
                break
                
            row = rows[0]
            store_name = get_prop_value(row, PR_DISPLAY_NAME) or "Unknown"
            is_default = get_prop_value(row, PR_DEFAULT_STORE)
            
            print(f"   - {store_name} {'⭐ (기본)' if is_default else ''}")
            
            # Default store인지 확인
            if is_default:
                store_eid = get_prop_value(row, PR_ENTRYID)
                
                # Message Store 열기
                store = session.OpenMsgStore(
                    0,
                    store_eid,
                    None,
                    mapi.MDB_WRITE | mapi.MAPI_DEFERRED_ERRORS
                )
                
                # Inbox 가져오기
                inbox_eid = store.GetReceiveFolder("IPM.Note", 0)[0]
                inbox = store.OpenEntry(inbox_eid, None, mapi.MAPI_MODIFY)
                
                print(f"   ✅ Inbox 열림: {store_name}")
                return store, inbox
                
        print("   ❌ Default store를 찾을 수 없음")
        return None, None
        
    except Exception as e:
        print(f"   ❌ Inbox 가져오기 실패: {e}")
        import traceback
        traceback.print_exc()
        return None, None

def process_messages(folder, folder_name="INBOX"):
    """메시지 처리"""
    print(f"\n📂 Processing: {folder_name}")
    
    try:
        # Contents 테이블 가져오기
        contents = folder.GetContentsTable(0)
        
        # 필요한 컬럼 설정
        contents.SetColumns((
            PR_ENTRYID,
            PR_SUBJECT,
            PR_MESSAGE_DELIVERY_TIME,
            PR_HASATTACH
        ), 0)
        
        count = 0
        total_messages = contents.GetRowCount(0)
        print(f"   📊 Total messages: {total_messages}")
        
        processed = 0
        while True:
            rows = contents.QueryRows(100, 0)
            if len(rows) == 0:
                break
                
            for row in rows:
                processed += 1
                if processed % 100 == 0:
                    print(f"   ⏳ 처리 중... {processed}/{total_messages}")
                
                try:
                    entry_id = get_prop_value(row, PR_ENTRYID)
                    subject = get_prop_value(row, PR_SUBJECT) or ""
                    delivery_time = get_prop_value(row, PR_MESSAGE_DELIVERY_TIME)
                    has_attach = get_prop_value(row, PR_HASATTACH)
                    
                    # 날짜 체크
                    if delivery_time:
                        msg_date = datetime.datetime.fromtimestamp(delivery_time.timestamp())
                        if msg_date < START_DATE or msg_date > END_DATE:
                            continue
                    else:
                        continue
                    
                    # 제목 체크
                    if TARGET_SUBJECT.lower() not in subject.lower():
                        continue
                    
                    # 첨부파일 체크
                    if not has_attach:
                        continue
                    
                    # 메시지 열기
                    message = folder.OpenEntry(entry_id, None, 0)
                    
                    # 첨부파일 테이블 가져오기
                    attach_table = message.GetAttachmentTable(0)
                    attach_table.SetColumns((PR_ATTACH_NUM, PR_ATTACH_FILENAME, PR_ATTACH_LONG_FILENAME), 0)
                    
                    while True:
                        attach_rows = attach_table.QueryRows(10, 0)
                        if len(attach_rows) == 0:
                            break
                            
                        for attach_row in attach_rows:
                            attach_num = get_prop_value(attach_row, PR_ATTACH_NUM)
                            filename = (get_prop_value(attach_row, PR_ATTACH_LONG_FILENAME) or 
                                      get_prop_value(attach_row, PR_ATTACH_FILENAME) or "")
                            
                            if filename.lower().endswith((".xlsx", ".xls")):
                                # 첨부파일 열기
                                attachment = message.OpenAttach(attach_num, None, 0)
                                
                                # 첨부파일 데이터 가져오기
                                stream = attachment.OpenProperty(
                                    PR_ATTACH_DATA_BIN,
                                    None,
                                    0,
                                    0
                                )
                                
                                # 데이터 읽기
                                data = b""
                                while True:
                                    chunk = stream.Read(4096)
                                    if not chunk:
                                        break
                                    data += chunk
                                
                                # 파일로 저장
                                final_name = get_custom_filename(RAW_FILES_FOLDER, msg_date)
                                save_path = os.path.join(RAW_FILES_FOLDER, final_name)
                                
                                with open(save_path, 'wb') as f:
                                    f.write(data)
                                
                                print(f"   ✅ Downloaded: {final_name}")
                                count += 1
                    
                except Exception as e:
                    continue
        
        print(f"   ✅ 완료: {count}개 파일 다운로드")
        return count
        
    except Exception as e:
        print(f"   ❌ 메시지 처리 오류: {e}")
        import traceback
        traceback.print_exc()
        return 0

def find_archive_folders(store):
    """Archive 폴더 찾기"""
    archives = []
    
    try:
        # Root folder 가져오기
        root_folder = store.OpenEntry(None, None, 0)
        
        # Hierarchy 테이블
        hierarchy = root_folder.GetHierarchyTable(mapi.CONVENIENT_DEPTH)
        hierarchy.SetColumns((PR_ENTRYID, PR_DISPLAY_NAME), 0)
        
        while True:
            rows = hierarchy.QueryRows(100, 0)
            if len(rows) == 0:
                break
                
            for row in rows:
                folder_name = get_prop_value(row, PR_DISPLAY_NAME) or ""
                
                # Archive 관련 폴더 찾기
                if any(kw in folder_name.lower() for kw in ["archive", "보관", "아카이브"]):
                    entry_id = get_prop_value(row, PR_ENTRYID)
                    folder = store.OpenEntry(entry_id, None, 0)
                    archives.append((folder, folder_name))
                    print(f"   ✅ Found Archive: {folder_name}")
        
    except Exception as e:
        print(f"   ⚠️ Archive 검색 오류: {e}")
    
    return archives

def step_1_download_emails():
    print("\n" + "=" * 60)
    print("--- [STEP 1] PURE MAPI EMAIL DOWNLOAD ---")
    print("=" * 60)
    
    os.makedirs(RAW_FILES_FOLDER, exist_ok=True)
    print(f"📂 Raw Files Folder: {RAW_FILES_FOLDER}")
    
    # MAPI 로그인
    session = mapi_login()
    if not session:
        return False
    
    try:
        # Inbox 가져오기
        store, inbox = get_inbox(session)
        if not inbox:
            return False
        
        # Inbox 처리
        total = process_messages(inbox, "INBOX")
        
        # Archive 폴더 검색 및 처리
        print("\n🔍 Searching for Archives...")
        archives = find_archive_folders(store)
        
        for archive_folder, archive_name in archives:
            total += process_messages(archive_folder, f"ARCHIVE: {archive_name}")
        
        print(f"\n🎉 Download Complete: {total} files downloaded.")
        
        # MAPI 정리
        try:
            session.Logoff(0, 0, 0)
            mapi.MAPIUninitialize()
        except:
            pass
        
        return True
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        import traceback
        traceback.print_exc()
        try:
            session.Logoff(0, 0, 0)
            mapi.MAPIUninitialize()
        except:
            pass
        return False

def step_2_process_excel_files():
    print("\n" + "=" * 60)
    print("--- [STEP 2] CLEANING & MERGING EXCEL FILES ---")
    print("=" * 60)
    
    master_path = os.path.join(OUTPUT_FOLDER, MASTER_FILENAME)
    search_pattern = os.path.join(RAW_FILES_FOLDER, "*Active-Applications*.xlsx")
    files = glob.glob(search_pattern)
    
    if not files:
        print("❌ No downloaded files found to process.")
        return

    print(f"ℹ️ Processing {len(files)} raw files...\n")
    
    df_list = []
    for f in files:
        try:
            df = pd.read_excel(f, header=2)
            df.columns = df.columns.str.strip()
            
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            df.dropna(how='all', axis=1, inplace=True)
            df.dropna(how='all', axis=0, inplace=True)
            if not df.empty:
                mask = df.iloc[:, 0].astype(str).str.contains("This page provides", na=False)
                df = df[~mask]

            df_list.append(df)
            print(f"   ✅ Processed: {os.path.basename(f)}")
        except Exception as e:
            print(f"   ⚠️ Error: {os.path.basename(f)}: {e}")

    if df_list:
        master_df = pd.concat(df_list, ignore_index=True)
        
        if 'Adm Appl Nbr' in master_df.columns:
            master_df.drop_duplicates(subset=['Adm Appl Nbr'], keep='last', inplace=True)
        
        if 'Prefered First Name' in master_df.columns and 'Prefered Last Name' in master_df.columns:
            master_df['Full_Name'] = (master_df['Prefered First Name'].fillna('') + " " + 
                                      master_df['Prefered Last Name'].fillna('')).str.lower().str.strip()

        master_df.to_excel(master_path, index=False, sheet_name="Active Applications")
        print(f"\n✅ Master File Created: {master_path}")
        print(f"   📊 Total Rows: {len(master_df)}")
    else:
        print("\n❌ No valid data to merge.")

if __name__ == "__main__":
    print("\n" + "🚀" * 30)
    print("PURE MAPI EMAIL DOWNLOADER")
    print("🚀" * 30)
    
    if step_1_download_emails():
        step_2_process_excel_files()
    
    print("\n" + "=" * 60)
    print("✅ SCRIPT COMPLETED")
    print("=" * 60)