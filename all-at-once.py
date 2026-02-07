import os
import datetime
import pandas as pd
import glob
import warnings
import win32com.client

# ================= CONFIGURATION =================
# Dynamic path - works on ANY computer automatically
USER_DOCS = os.path.join(os.environ["USERPROFILE"], "Documents")
OUTPUT_FOLDER = os.path.join(USER_DOCS, "student-applications")
RAW_FILES_FOLDER = os.path.join(OUTPUT_FOLDER, "files")
MASTER_FILENAME = "Active Application Details.xlsx"

# Search Settings
TARGET_SUBJECT = "NSN Report"
START_DATE = datetime.datetime(2025, 9, 1)
END_DATE = datetime.datetime(2026, 2, 9, 23, 59, 59)

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

# ================= STEP 1: DOWNLOAD (Using the method that works for you) =================

def process_outlook_folder(folder, label: str) -> int:
    print(f"\n📂 Scanning: {label}")
    
    try:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        
        # Fast Filtering
        start_str = START_DATE.strftime("%m/%d/%Y 00:00 AM")
        end_str   = END_DATE.strftime("%m/%d/%Y 11:59 PM")
        filter_str = f"[ReceivedTime] >= '{start_str}' AND [ReceivedTime] <= '{end_str}'"
        
        items = items.Restrict(filter_str)
        print(f"   📊 Found {items.Count} items in date range.")
    except:
        print("   ⚠️ Filter failed. Scanning all items (slower)...")
    
    count = 0
    for msg in items:
        try:
            # Double check date (Python side)
            msg_time = msg.ReceivedTime
            # Convert to naive datetime for comparison
            msg_dt = datetime.datetime(msg_time.year, msg_time.month, msg_time.day, msg_time.hour, msg_time.minute, msg_time.second)
            
            if not (START_DATE <= msg_dt <= END_DATE):
                continue

            # Check Subject
            if TARGET_SUBJECT.lower() not in (msg.Subject or "").lower():
                continue

            # Check Attachments
            if msg.Attachments.Count > 0:
                for att in msg.Attachments:
                    fname = att.FileName or ""
                    if fname.lower().endswith((".xlsx", ".xls")):
                        final_name = get_custom_filename(RAW_FILES_FOLDER, msg_dt)
                        save_path = os.path.join(RAW_FILES_FOLDER, final_name)
                        
                        att.SaveAsFile(save_path)
                        print(f"   ✅ Downloaded: {final_name}")
                        count += 1
        except Exception:
            continue
            
    return count

def step_1_download_emails():
    print("\n" + "=" * 60)
    print("--- [STEP 1] DOWNLOADING EMAILS (Active App Mode) ---")
    print("=" * 60)
    
    if not os.path.exists(RAW_FILES_FOLDER):
        os.makedirs(RAW_FILES_FOLDER)
        print(f"📁 Created folder: {RAW_FILES_FOLDER}")

    try:
        # Connect to the OPEN Outlook app
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except:
        print("❌ Could not connect to Outlook. Is it open?")
        return False

    inbox = outlook.GetDefaultFolder(6) # 6 = Inbox
    total = process_outlook_folder(inbox, "INBOX")

    # Optional: Check Archive (Basic check)
    try:
        root = inbox.Parent
        for folder in root.Folders:
            if "archive" in folder.Name.lower():
                total += process_outlook_folder(folder, f"ARCHIVE: {folder.Name}")
    except:
        pass

    print(f"\n🎉 Step 1 Complete: {total} files downloaded.")
    return total > 0 # Return True if files were downloaded (or just execute Step 2 anyway)

# ================= STEP 2: PROCESS (Using Pandas) =================

def step_2_process_excel_files():
    print("\n" + "=" * 60)
    print("--- [STEP 2] CLEANING & MERGING DATA ---")
    print("=" * 60)
    
    master_path = os.path.join(OUTPUT_FOLDER, MASTER_FILENAME)
    search_pattern = os.path.join(RAW_FILES_FOLDER, "*Active-Applications*.xlsx")
    files = glob.glob(search_pattern)
    
    if not files:
        print("❌ No files found to process.")
        return

    print(f"ℹ️ Processing {len(files)} raw files...\n")
    
    df_list = []
    for f in files:
        try:
            # Load and Clean
            df = pd.read_excel(f, header=2)
            df.columns = df.columns.str.strip()
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            df.dropna(how='all', axis=1, inplace=True)
            df.dropna(how='all', axis=0, inplace=True)
            
            if not df.empty:
                # Remove header repetition inside data
                mask = df.iloc[:, 0].astype(str).str.contains("This page provides", na=False)
                df = df[~mask]
                df_list.append(df)
                
            print(f"   ✅ Merged: {os.path.basename(f)}")
        except Exception as e:
            print(f"   ⚠️ Error reading {os.path.basename(f)}: {e}")

    if df_list:
        master_df = pd.concat(df_list, ignore_index=True)
        
        # Remove duplicates based on App Number
        if 'Adm Appl Nbr' in master_df.columns:
            master_df.drop_duplicates(subset=['Adm Appl Nbr'], keep='last', inplace=True)
        
        # Create Full Name
        if 'Prefered First Name' in master_df.columns and 'Prefered Last Name' in master_df.columns:
            master_df['Full_Name'] = (master_df['Prefered First Name'].fillna('') + " " + 
                                      master_df['Prefered Last Name'].fillna('')).str.lower().str.strip()

        master_df.to_excel(master_path, index=False, sheet_name="Active Applications")
        print(f"\n✅ SUCCESS! Master File Created: {master_path}")
        print(f"   📊 Total Unique Rows: {len(master_df)}")
        
        # Ask to open
        try:
            os.startfile(master_path)
        except:
            pass
    else:
        print("\n❌ No valid data extracted.")

# ================= MAIN EXECUTION =================

if __name__ == "__main__":
    print("🚀 STARTING AUTOMATION...")
    
    # Run Step 1
    step_1_download_emails()
    
    # Always run Step 2 (to process existing files if any)
    step_2_process_excel_files()
    
    print("\n✅ PROCESS FINISHED")