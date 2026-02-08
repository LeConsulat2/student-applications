import os
import datetime
import win32com.client
import pandas as pd
import glob
import warnings

# ================= CONFIGURATION =================
# 1. Folder Settings (Separated for Output vs Input)
# Automatically finds "C:\Users\User\Documents"
user_docs = os.path.join(os.environ["USERPROFILE"], "Documents")

# Where the FINAL merged report will be saved
OUTPUT_FOLDER = os.path.join(user_docs, "student-applications")

# Where the RAW downloaded files will go
RAW_FILES_FOLDER = os.path.join(OUTPUT_FOLDER, "files")

MASTER_FILENAME = "Active Application Details.xlsx"
TARGET_SUBJECT = "NSN Report"

# 2. Date Settings
START_DATE = datetime.datetime(2025, 9, 1)
END_DATE = datetime.datetime(2026, 2, 5, 23, 59, 59)

# 3. Suppress Warnings
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

def find_folders_recursive(folder, keywords, found=None):
    if found is None: found = []
    try:
        path = folder.FolderPath or ""
        for kw in keywords:
            if kw.lower() in path.lower():
                found.append(folder)
                break 
        for i in range(1, folder.Folders.Count + 1):
            find_folders_recursive(folder.Folders.Item(i), keywords, found)
    except Exception:
        pass
    return found

def process_folder(folder, label: str) -> int:
    print(f"\n📂 Scanning: {label}")
    try: print(f"   - Name: {folder.Name}")
    except: pass

    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    
    start_str = START_DATE.strftime("%m/%d/%Y 00:00 AM")
    end_str   = END_DATE.strftime("%m/%d/%Y 11:59 PM")
    filter_str = f"[ReceivedTime] >= '{start_str}' AND [ReceivedTime] <= '{end_str}'"
    
    try:
        items_to_process = items.Restrict(filter_str)
        print(f"   📊 Found {items_to_process.Count} items in date range")
    except:
        items_to_process = items

    count = 0
    for msg in items_to_process:
        try:
            received = msg.ReceivedTime
            received_dt = datetime.datetime(received.year, received.month, received.day, 
                                          received.hour, received.minute, received.second)
            
            if received_dt < START_DATE or received_dt > END_DATE: continue
            if TARGET_SUBJECT.lower() not in (msg.Subject or "").lower(): continue

            atts = msg.Attachments
            if atts.Count == 0: continue

            for i in range(1, atts.Count + 1):
                att = atts.Item(i)
                fname = att.FileName or ""
                if fname.lower().endswith((".xlsx", ".xls")):
                    # Save to the RAW files folder
                    final_name = get_custom_filename(RAW_FILES_FOLDER, received_dt)
                    save_path = os.path.join(RAW_FILES_FOLDER, final_name)
                    att.SaveAsFile(save_path)
                    print(f"   ✅ Downloaded: {final_name}")
                    count += 1
        except Exception:
            continue
    return count

def step_1_download_emails():
    print("--- [STEP 1] DOWNLOADING FROM OUTLOOK APP ---")
    
    # Create the raw files folder if it doesn't exist
    os.makedirs(RAW_FILES_FOLDER, exist_ok=True)
    print(f"📂 Saving Raw Files to: {RAW_FILES_FOLDER}")
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        print(f"\n❌ CRITICAL ERROR: Cannot connect to Outlook.")
        print("   👉 Solution: Please open 'Classic Outlook' and try again.")
        return False

    inbox = outlook.GetDefaultFolder(6)
    root = inbox.Parent
    primary_name = root.Name
    print(f"👤 Account: {primary_name}")
    
    total = process_folder(inbox, "INBOX")

    print("\n🔍 Searching for Archives...")
    sub_archives = find_folders_recursive(root, ["\\archive", "\\보관", "\\아카이브"])
    
    for i in range(1, outlook.Stores.Count + 1):
        try:
            store = outlook.Stores.Item(i)
            name = store.DisplayName
            if "archive" in name.lower() and primary_name in name:
                print(f"   ✅ Found Archive Store: {name}")
                store_root = store.GetRootFolder()
                sub_archives.extend(find_folders_recursive(store_root, ["root", "top", "archive", "inbox"]))
        except: pass

    unique_archives = {f.FolderPath: f for f in sub_archives}
    if unique_archives:
        for path, folder in unique_archives.items():
            total += process_folder(folder, "ARCHIVE")
            
    print(f"🎉 Download Complete: {total} files.\n")
    return True

def step_2_process_excel_files():
    print("--- [STEP 2] CLEANING & MERGING EXCEL FILES ---")
    
    # OUTPUT: Save the Master file one level UP (in OUTPUT_FOLDER)
    master_path = os.path.join(OUTPUT_FOLDER, MASTER_FILENAME)
    
    # INPUT: Read files from the 'files' folder (RAW_FILES_FOLDER)
    search_pattern = os.path.join(RAW_FILES_FOLDER, "*Active-Applications*.xlsx")
    files = glob.glob(search_pattern)
    
    if not files:
        print("❌ No downloaded files found to process.")
        return

    print(f"ℹ️ Processing {len(files)} raw files from '{os.path.basename(RAW_FILES_FOLDER)}' folder...")
    
    df_list = []
    for f in files:
        try:
            df = pd.read_excel(f, header=2)
            df.columns = df.columns.str.strip()
            
            # --- CLEANING ---
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
            df.dropna(how='all', axis=1, inplace=True)
            df.dropna(how='all', axis=0, inplace=True)
            if not df.empty:
                mask = df.iloc[:, 0].astype(str).str.contains("This page provides", na=False)
                df = df[~mask]

            df_list.append(df)
        except Exception as e:
            print(f"⚠️ Error reading {os.path.basename(f)}: {e}")

    if df_list:
        master_df = pd.concat(df_list, ignore_index=True)
        
        if 'Adm Appl Nbr' in master_df.columns:
            master_df.drop_duplicates(subset=['Adm Appl Nbr'], keep='last', inplace=True)
        
        if 'Prefered First Name' in master_df.columns and 'Prefered Last Name' in master_df.columns:
            master_df['Full_Name'] = (master_df['Prefered First Name'].fillna('') + " " + 
                                      master_df['Prefered Last Name'].fillna('')).str.lower().str.strip()

        # Save to the main folder
        master_df.to_excel(master_path, index=False, sheet_name="Active Applications")
        print(f"✅ [SUCCESS] Master File Created: {master_path}")
        print(f"   Total Rows: {len(master_df)}")
    else:
        print("❌ No valid data found to merge.")

if __name__ == "__main__":
    if step_1_download_emails():
        step_2_process_excel_files()
    
    print("\n✅ ALL TASKS COMPLETED.")