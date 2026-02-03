import os
import datetime
import win32com.client

# ================= CONFIGURATION =================
SAVE_FOLDER = r"C:\Users\wooin\Documents\student-applications\files"
TARGET_SUBJECT = "NSN Report"

START_DATE = datetime.datetime(2025, 9, 1)
END_DATE   = datetime.datetime(2026, 2, 4, 23, 59, 59)
# =================================================

def get_custom_filename(folder: str, date_obj: datetime.datetime) -> str:
    """
    Creates a filename that SORTS correctly but READS like NZ format.
    Format: YYYY-MM-DD__Active-Applications-001__(DD-MM-YYYY).xlsx
    """
    # 1. Sortable Prefix (Year-Month-Day) -> Computer uses this to sort
    sort_prefix = date_obj.strftime("%Y-%m-%d")
    
    # 2. Readable Suffix (Day-Month-Year) -> For you to read
    nz_suffix = date_obj.strftime("%d-%m-%Y")

    counter = 1
    while True:
        # Example: 2025-12-04__Active-Applications-001__(04-12-2025).xlsx
        filename = f"{sort_prefix}__Active-Applications-{counter:03d}__({nz_suffix}).xlsx"
        full_path = os.path.join(folder, filename)
        
        if not os.path.exists(full_path):
            return filename
        counter += 1

def find_folders_recursive(folder, keywords, found=None):
    if found is None:
        found = []
    try:
        path = folder.FolderPath or ""
        # Check current folder
        for kw in keywords:
            if kw.lower() in path.lower():
                found.append(folder)
                break 
        # Check subfolders
        for i in range(1, folder.Folders.Count + 1):
            find_folders_recursive(folder.Folders.Item(i), keywords, found)
    except Exception:
        pass
    return found

def process_folder(folder, label: str) -> int:
    print(f"\n📂 Scanning: {label}")
    try:
        print(f"   - Name: {folder.Name}")
    except:
        pass

    items = folder.Items
    items.Sort("[ReceivedTime]", True)
    
    # --- FAST FILTERING (Restrict) ---
    start_str = START_DATE.strftime("%m/%d/%Y 00:00 AM")
    end_str   = END_DATE.strftime("%m/%d/%Y 11:59 PM")
    filter_str = f"[ReceivedTime] >= '{start_str}' AND [ReceivedTime] <= '{end_str}'"
    
    try:
        items_to_process = items.Restrict(filter_str)
        print(f"   📊 Found {items_to_process.Count} items in date range")
    except Exception:
        print("   ⚠️ Filter error, checking items manually...")
        items_to_process = items

    count = 0
    for msg in items_to_process:
        try:
            # Double check date
            received = msg.ReceivedTime
            received_dt = datetime.datetime(
                received.year, received.month, received.day,
                received.hour, received.minute, received.second
            )
            
            if received_dt < START_DATE or received_dt > END_DATE:
                continue

            if TARGET_SUBJECT.lower() not in (msg.Subject or "").lower():
                continue

            atts = msg.Attachments
            if atts.Count == 0:
                continue

            # --- PROCESS ATTACHMENTS ---
            for i in range(1, atts.Count + 1):
                att = atts.Item(i)
                fname = att.FileName or ""
                
                if fname.lower().endswith((".xlsx", ".xls")):
                    # Pass the DATE OBJECT, not a string
                    final_name = get_custom_filename(SAVE_FOLDER, received_dt)
                    save_path = os.path.join(SAVE_FOLDER, final_name)
                    
                    att.SaveAsFile(save_path)
                    print(f"   ✅ Downloaded: {final_name}")
                    count += 1

        except Exception:
            continue

    return count

def download_safe_mode():
    os.makedirs(SAVE_FOLDER, exist_ok=True)
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except:
        print("❌ Could not connect to Outlook.")
        return

    # 1. Get Primary Inbox
    inbox = outlook.GetDefaultFolder(6) 
    root = inbox.Parent 
    primary_name = root.Name 
    
    print(f"👤 Primary Account: {primary_name}")
    print(f"📅 Date range: {START_DATE.date()} ~ {END_DATE.date()}")
    
    total = 0

    # 2. Process Inbox
    total += process_folder(inbox, "INBOX")

    # 3. Find Archives SAFELY (Only yours)
    print("\n🔍 Searching for your Archives...")
    
    # A) Internal Archives
    sub_archives = find_folders_recursive(root, ["\\archive", "\\보관", "\\아카이브"])
    
    # B) Online Archives (Verify ownership)
    for i in range(1, outlook.Stores.Count + 1):
        try:
            store = outlook.Stores.Item(i)
            name = store.DisplayName
            # MUST contain "Archive" AND your name
            if "archive" in name.lower() and primary_name in name:
                print(f"   ✅ Found User Archive Store: {name}")
                store_root = store.GetRootFolder()
                sub_archives.extend(find_folders_recursive(store_root, ["root", "top", "archive", "inbox"]))
        except:
            pass

    # Process Archives
    unique_archives = {f.FolderPath: f for f in sub_archives}
    if unique_archives:
        for path, folder in unique_archives.items():
            total += process_folder(folder, "ARCHIVE")

    print(f"\n🎉 Done. Downloaded {total} file(s).")

if __name__ == "__main__":
    download_safe_mode()