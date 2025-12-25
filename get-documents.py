import os
import datetime
import win32com.client

# ================= 설정값 =================
SAVE_FOLDER = r"C:\Users\wooin\Documents\student-applications\files"
TARGET_SUBJECT = "NSN Report"

START_DATE = datetime.datetime(2025, 9, 1)
END_DATE   = datetime.datetime(2025, 12, 31, 23, 59, 59)
# ==========================================

def get_custom_filename(folder: str, date_str: str) -> str:
    """
    파일명 형식: Active-Applications-001-01-09-2025.xlsx
    폴더 내 모든 파일 중 가장 큰 번호를 찾아서 +1
    같은 날짜든 다른 날짜든 상관없이 계속 증가
    """
    # 폴더 내 모든 Active-Applications 파일의 번호를 찾기
    max_counter = 0
    try:
        for existing_file in os.listdir(folder):
            if existing_file.startswith("Active-Applications-") and existing_file.endswith(".xlsx"):
                # Active-Applications-001-01-09-2025.xlsx에서 001 추출
                parts = existing_file.split("-")
                if len(parts) >= 3:
                    try:
                        counter = int(parts[2])  # 001, 002, 003...
                        max_counter = max(max_counter, counter)
                    except ValueError:
                        continue
    except Exception:
        pass
    
    # 다음 번호 사용
    next_counter = max_counter + 1
    filename = f"Active-Applications-{next_counter:03d}-{date_str}.xlsx"
    return filename


def find_folders_recursive(folder, keywords, found=None):
    if found is None:
        found = []

    try:
        path = folder.FolderPath or ""
        low = path.lower()
        for kw in keywords:
            if kw.lower() in low:
                found.append(folder)
                break
    except Exception:
        pass

    for i in range(1, folder.Folders.Count + 1):
        find_folders_recursive(folder.Folders.Item(i), keywords, found)

    return found


# ============================================================
# ✅✅✅ FIXED: Use proper Outlook date format for Restrict
# ============================================================
def restrict_by_date(items, start_dt, end_dt):
    """
    Outlook Restrict requires specific date format.
    Using ISO-like format: "MM/DD/YYYY HH:MM AM/PM"
    But we need to convert 24h to 12h format properly
    """
    # Convert to 12-hour format properly
    start_str = start_dt.strftime("%m/%d/%Y 12:00 AM")
    end_str = end_dt.strftime("%m/%d/%Y 11:59 PM")
    
    filter_str = f"[ReceivedTime] >= '{start_str}' AND [ReceivedTime] <= '{end_str}'"
    
    print(f"   🔍 Filter: {filter_str}")
    
    try:
        restricted = items.Restrict(filter_str)
        return restricted
    except Exception as e:
        print(f"   ⚠️ Restrict failed: {e}")
        print("   ⚠️ Falling back to manual filtering")
        return None


def process_folder(folder, label: str) -> int:
    count = 0

    print(f"\n📂 Scanning: {label}")
    try:
        print(f"   - Name: {folder.Name}")
        print(f"   - Path: {folder.FolderPath}")
    except Exception:
        pass

    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # Most recent first
    
    # Try Restrict first
    restricted_items = restrict_by_date(items, START_DATE, END_DATE)
    
    # If Restrict failed, use all items and filter manually
    if restricted_items is None:
        print("   ⚠️ Using manual date filtering (slower but more reliable)")
        items_to_process = items
        use_manual_filter = True
    else:
        items_to_process = restricted_items
        use_manual_filter = False
        try:
            print(f"   📊 Restrict returned {items_to_process.Count} items")
        except:
            pass

    processed_count = 0
    for msg in items_to_process:
        try:
            # Manual date filter if Restrict failed
            if use_manual_filter:
                received = msg.ReceivedTime
                received_dt = datetime.datetime(
                    received.year, received.month, received.day,
                    received.hour, received.minute, received.second
                )
                
                if received_dt < START_DATE or received_dt > END_DATE:
                    continue
            else:
                received = msg.ReceivedTime
                received_dt = datetime.datetime(
                    received.year, received.month, received.day,
                    received.hour, received.minute, received.second
                )
            
            # Progress indicator
            processed_count += 1
            if processed_count % 100 == 0:
                print(f"   ... processed {processed_count} items")
            
            subject = (msg.Subject or "")
            if TARGET_SUBJECT.lower() not in subject.lower():
                continue

            atts = msg.Attachments
            if atts.Count == 0:
                continue

            date_str = received_dt.strftime("%d-%m-%Y")

            for i in range(1, atts.Count + 1):
                att = atts.Item(i)
                fname = att.FileName or ""
                if fname.lower().endswith((".xlsx", ".xls")):
                    final_name = get_custom_filename(SAVE_FOLDER, date_str)
                    save_path = os.path.join(SAVE_FOLDER, final_name)
                    att.SaveAsFile(save_path)
                    print(f"   ✅ Downloaded: {final_name} (Received: {received_dt})")
                    count += 1

        except Exception as e:
            # More verbose error logging
            try:
                print(f"   ❌ Error processing email: {msg.Subject[:50]} - {str(e)}")
            except:
                print(f"   ❌ Error processing email: {str(e)}")
            continue

    print(f"   📊 Processed {processed_count} total items in this folder")
    return count


def download_everything_in_inbox_and_archives():
    os.makedirs(SAVE_FOLDER, exist_ok=True)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    inbox = outlook.GetDefaultFolder(6)
    root = inbox.Parent  # mailbox root

    print(f"📅 Date range: {START_DATE} ~ {END_DATE}")
    print(f"💾 Save folder: {SAVE_FOLDER}")
    print(f"🎯 Target subject: {TARGET_SUBJECT}")

    total = 0

    # 1) INBOX
    total += process_folder(inbox, "INBOX")

    # 2) ARCHIVE folders
    archive_keywords = ["\\archive", "online archive", "\\보관", "\\아카이브", "in-place archive"]
    archive_folders = find_folders_recursive(root, archive_keywords)

    # Remove duplicates
    unique = {}
    for f in archive_folders:
        try:
            unique[f.FolderPath] = f
        except Exception:
            pass
    archive_folders = list(unique.values())

    if not archive_folders:
        print("\n⚠️ No archive-like folders found under this mailbox root.")
        print("   -> Checking all Stores for Online Archive...")
        
        # Check all stores (including Online Archive)
        for store_idx in range(1, outlook.Stores.Count + 1):
            store = outlook.Stores.Item(store_idx)
            try:
                print(f"\n📦 Store: {store.DisplayName}")
                store_root = store.GetRootFolder()
                
                # Check if this is an archive store
                if any(kw in store.DisplayName.lower() for kw in ["archive", "보관", "아카이브"]):
                    total += process_folder(store_root, f"ARCHIVE_STORE[{store.DisplayName}]")
                    
                    # Also check subfolders
                    for folder_idx in range(1, store_root.Folders.Count + 1):
                        subfolder = store_root.Folders.Item(folder_idx)
                        total += process_folder(subfolder, f"ARCHIVE_STORE_SUB[{subfolder.Name}]")
            except Exception as e:
                print(f"   ❌ Error accessing store: {e}")
    else:
        print(f"\n✅ Found {len(archive_folders)} archive-like folder(s).")
        for f in archive_folders:
            try:
                print("   -", f.FolderPath)
            except Exception:
                pass

        for idx, f in enumerate(archive_folders, start=1):
            total += process_folder(f, f"ARCHIVE[{idx}]")

    print(f"\n🎉 Done. Downloaded {total} file(s).")


if __name__ == "__main__":
    download_everything_in_inbox_and_archives()