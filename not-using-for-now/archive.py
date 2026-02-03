import os
import datetime
import win32com.client

SAVE_FOLDER = r"C:\Users\wooin\Documents\student-applications\files"
TARGET_SUBJECT = "NSN Report"
START_DATE = datetime.datetime(2025, 9, 1)
END_DATE = datetime.datetime(2025, 12, 31, 23, 59, 59)

def get_unique_filename(folder, filename):
    base, ext = os.path.splitext(filename)
    counter = 1
    new_filename = filename
    while os.path.exists(os.path.join(folder, new_filename)):
        new_filename = f"{base}({counter}){ext}"
        counter += 1
    return new_filename

def find_folder_by_names(root_folder, candidate_names):
    """Outlook 폴더 트리에서 이름이 candidate_names 중 하나인 폴더를 찾아 반환"""
    cand = {n.lower() for n in candidate_names}
    # root 바로 아래부터 훑기
    for i in range(1, root_folder.Folders.Count + 1):
        f = root_folder.Folders.Item(i)
        if (f.Name or "").lower() in cand:
            return f
    return None

def process_folder(folder, label):
    items = folder.Items
    items.Sort("[ReceivedTime]", True)

    count = 0
    print(f"\n📂 Scanning: {label} ({folder.Name})")

    for msg in items:
        try:
            received = msg.ReceivedTime
            received_dt = datetime.datetime(
                received.year, received.month, received.day,
                received.hour, received.minute, received.second
            )

            if received_dt > END_DATE:
                continue
            if received_dt < START_DATE:
                break  # 최신순이므로 여기서 종료

            subject = (msg.Subject or "")
            if TARGET_SUBJECT.lower() not in subject.lower():
                continue

            atts = msg.Attachments
            if atts.Count == 0:
                continue

            for i in range(1, atts.Count + 1):
                att = atts.Item(i)
                fname = att.FileName or ""
                if fname.lower().endswith((".xlsx", ".xls")):
                    date_str = received_dt.strftime("%d-%m-%Y")
                    new_name = f"{date_str}_{fname}"
                    final_name = get_unique_filename(SAVE_FOLDER, new_name)
                    save_path = os.path.join(SAVE_FOLDER, final_name)

                    att.SaveAsFile(save_path)
                    print(f"✅ Downloaded: {final_name}")
                    count += 1

        except Exception:
            continue

    return count

def download_inbox_and_archive():
    os.makedirs(SAVE_FOLDER, exist_ok=True)

    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Inbox
    inbox = outlook.GetDefaultFolder(6)

    # Archive 폴더는 환경마다 이름이 다를 수 있음 → 후보 이름들로 탐색
    # (예: "Archive", "In-Place Archive", "Online Archive", "보관", "아카이브" 등)
    root = inbox.Parent  # 보통 같은 mailbox 루트
    archive = find_folder_by_names(root, [
        "Archive", "In-Place Archive", "Online Archive", "보관", "아카이브"
    ])

    print(f"📁 Saving to: {SAVE_FOLDER}")
    total = 0
    total += process_folder(inbox, "INBOX")

    if archive:
        total += process_folder(archive, "ARCHIVE")
    else:
        print("\n⚠️ Archive 폴더를 이름으로 찾지 못했음.")
        print("→ 네 Outlook에서 Archive 폴더의 정확한 이름을 확인해서 후보에 추가하면 됨.")

    print(f"\n🎉 Done. Downloaded {total} file(s).")

if __name__ == "__main__":
    download_inbox_and_archive()
