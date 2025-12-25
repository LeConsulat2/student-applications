import os
import datetime
import win32com.client
from datetime import datetime as dt

# ================= CONFIGURATION =================
SAVE_FOLDER = r"C:\Users\wooin\Documents\student-applications\files"
TARGET_SUBJECT_KEYWORD = "NSN Report"

# Set the window: Sept 1st to Dec 31st, 2025
START_DATE = dt(2025, 9, 1, 0, 0, 0)
END_DATE   = dt(2025, 12, 31, 23, 59, 59)

# Folders to ignore to speed up the process
IGNORE_FOLDERS = ["Deleted Items", "Junk Email", "Drafts", "Outbox", "Sync Issues", "Calendar", "Contacts", "Tasks"]
# =================================================

def get_next_sequence_number(folder):
    """Calculates the next 000-style index based on existing files."""
    max_counter = 0
    if not os.path.exists(folder):
        return 1
    for f in os.listdir(folder):
        if f.startswith("Active-Applications-") and f.endswith(".xlsx"):
            try:
                # Extracts the number from 'Active-Applications-XXX-...'
                parts = f.split("-")
                if len(parts) >= 3:
                    max_counter = max(max_counter, int(parts[2]))
            except (ValueError, IndexError):
                continue
    return max_counter + 1

def process_folder_recursive(folder):
    """Deep-crawls every subfolder in the given folder."""
    download_count = 0
    
    # 1. Skip ignored system folders
    if folder.Name in IGNORE_FOLDERS:
        return 0

    print(f"🔍 Checking: {folder.FolderPath}")

    # 2. Get items and sort
    try:
        items = folder.Items
        # We don't use Restrict() here to avoid regional date format bugs
        # Instead, we sort and then filter manually for 100% accuracy
        items.Sort("[ReceivedTime]", True)
    except Exception as e:
        print(f"   ⚠️ Could not access items in {folder.Name}: {e}")
        return 0

    # 3. Iterate through emails
    for msg in items:
        try:
            # Basic check: Is it an email?
            if msg.Class != 43: # 43 = olMail
                continue

            # Robust Date Check
            # outlook_date is a timezone-aware object; we convert to naive for comparison
            received_time = msg.ReceivedTime
            msg_date = dt(received_time.year, received_time.month, received_time.day, 
                          received_time.hour, received_time.minute, received_time.second)

            # Optimization: If the email is older than our start date, and we are sorted 
            # by newest first, we can stop checking this folder's items.
            if msg_date < START_DATE:
                break
            
            # Skip if newer than our end range
            if msg_date > END_DATE:
                continue

            # Keyword Match (Subject or Body for safety)
            subject = (msg.Subject or "")
            if TARGET_SUBJECT_KEYWORD.lower() not in subject.lower():
                continue

            # Process Attachments
            atts = msg.Attachments
            if atts.Count > 0:
                date_str = msg_date.strftime("%d-%m-%Y")
                for i in range(1, atts.Count + 1):
                    att = atts.Item(i)
                    if att.FileName.lower().endswith((".xlsx", ".xls")):
                        seq = get_next_sequence_number(SAVE_FOLDER)
                        filename = f"Active-Applications-{seq:03d}-{date_str}.xlsx"
                        save_path = os.path.join(SAVE_FOLDER, filename)
                        
                        att.SaveAsFile(save_path)
                        print(f"   ✅ Saved: {filename} (from {folder.Name})")
                        download_count += 1

        except Exception as e:
            continue # Skip individual item errors (e.g. encrypted emails)

    # 4. Recursively check subfolders
    for i in range(1, folder.Folders.Count + 1):
        try:
            download_count += process_folder_recursive(folder.Folders.Item(i))
        except:
            continue

    return download_count

def main():
    print("🚀 Starting Professional NSN Recovery Script...")
    os.makedirs(SAVE_FOLDER, exist_ok=True)
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        print(f"❌ Error: Could not connect to Outlook. {e}")
        return

    total_files = 0

    # This is the "Sophisticated" part: 
    # We don't just look at the Inbox; we look at EVERY Store (Online Archive, Primary, etc.)
    for store in outlook.Stores:
        print(f"\n📦 Accessing Store: {store.DisplayName}")
        try:
            root_folder = store.GetRootFolder()
            total_files += process_folder_recursive(root_folder)
        except Exception as e:
            print(f"   ⚠️ Could not open Store {store.DisplayName}: {e}")

    print("-" * 30)
    print(f"🎉 SUCCESS: Captured {total_files} total files from Sept to Dec.")
    print(f"📂 Location: {SAVE_FOLDER}")

if __name__ == "__main__":
    main()