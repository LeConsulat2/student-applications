import os
import datetime
import win32com.client

# ================= 설정값 =================
SAVE_FOLDER = r"C:\Users\wooin\Documents\student-applications\files"
TARGET_SUBJECT = "NSN Report"

START_DATE = datetime.datetime(2025, 9, 1)
END_DATE = datetime.datetime(2025, 12, 31, 23, 59, 59)
# ==========================================


def get_custom_filename(folder: str, date_str: str) -> str:
    """
    ✅ 커스텀 파일명 형식:
        Active-Applications-001-01-09-2025.xlsx
        Active-Applications-002-01-09-2025.xlsx
    같은 날짜(date_str)에 이미 파일이 있으면 001 -> 002 -> 003 으로 증가
    """
    counter = 1
    while True:
        filename = f"Active-Applications-{counter:03d}-{date_str}.xlsx"
        full_path = os.path.join(folder, filename)
        if not os.path.exists(full_path):
            return filename
        counter += 1


# ============================================================
# ✅✅✅ [핵심 수정 포인트] Archive 폴더를 "이름"이 아니라
# FolderPath를 기준으로 재귀 탐색해서 찾는다.
#
# 이유:
# - UoA/O365 환경에서 Archive는 종종 "Online Archive - 이름" 같은
#   구조로 들어가서, 이름 매칭(find_folder_by_names)로는 엉뚱한 폴더를
#   잡거나 아예 못 잡는 경우가 많음.
# - 그래서 FolderPath에 'archive', 'online archive', '보관' 등이 포함된
#   폴더를 트리 전체에서 재귀로 찾아야 안정적임.
# ============================================================
def find_folder_recursive(folder, keywords):
    """
    Outlook 폴더 트리를 재귀로 순회하면서
    FolderPath에 keywords 중 하나가 포함된 폴더를 반환.

    keywords 예:
      ["\\archive", "online archive", "\\보관", "\\아카이브", "in-place archive"]
    """
    try:
        path = folder.FolderPath or ""
        for kw in keywords:
            if kw.lower() in path.lower():
                return folder
    except Exception:
        pass

    for i in range(1, folder.Folders.Count + 1):
        found = find_folder_recursive(folder.Folders.Item(i), keywords)
        if found:
            return found

    return None


def process_folder(folder, label: str) -> int:
    """
    특정 Outlook 폴더(Inbox/Archive 등)에서
    - 날짜 범위
    - 제목 포함 여부
    - 첨부 엑셀(.xlsx/.xls)
    조건에 맞는 첨부파일을 SAVE_FOLDER에 저장
    """
    items = folder.Items
    items.Sort("[ReceivedTime]", True)  # 최신 메일 순서

    count = 0
    print(f"\n📂 검색 시작: {label} (폴더명: {folder.Name})")
    try:
        print(f"   🔎 FolderPath: {folder.FolderPath}")  # 디버그용 (진짜 Archive 잡았는지 확인)
    except Exception:
        pass

    for msg in items:
        try:
            received = msg.ReceivedTime
            received_dt = datetime.datetime(
                received.year, received.month, received.day,
                received.hour, received.minute, received.second
            )

            # 날짜 범위 체크
            if received_dt > END_DATE:
                continue
            if received_dt < START_DATE:
                break  # 최신순 정렬이므로 여기서 끊어도 됨

            # 제목 체크
            subject = (msg.Subject or "")
            if TARGET_SUBJECT.lower() not in subject.lower():
                continue

            # 첨부 체크
            atts = msg.Attachments
            if atts.Count == 0:
                continue

            for i in range(1, atts.Count + 1):
                att = atts.Item(i)
                fname = att.FileName or ""

                # 엑셀 파일만 저장
                if fname.lower().endswith((".xlsx", ".xls")):
                    # ✅ 날짜는 네가 원한 NZ 표시용: DD-MM-YYYY
                    date_str = received_dt.strftime("%d-%m-%Y")

                    # ✅ 커스텀 파일명 생성 (001/002/003 + 날짜)
                    final_name = get_custom_filename(SAVE_FOLDER, date_str)
                    save_path = os.path.join(SAVE_FOLDER, final_name)

                    att.SaveAsFile(save_path)
                    print(f"   ✅ 다운로드 완료: {final_name}")
                    count += 1

        except Exception:
            # Outlook 아이템 중 간혹 에러나는 메시지/형식이 있음 → 스킵
            continue

    return count


def download_inbox_and_archive():
    os.makedirs(SAVE_FOLDER, exist_ok=True)

    # Outlook 세션 연결
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        print("❌ Outlook 연결 실패: Outlook 프로그램 설치/로그인 상태 확인 필요")
        print("에러:", e)
        return

    # 1) Inbox
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox

    # ============================================================
    # ✅✅✅ [핵심 수정 포인트] Archive 찾기
    # 기존: find_folder_by_names(root, candidates)  ← 이름 매칭이라 불안정
    # 변경: find_folder_recursive(root, keywords)   ← FolderPath 기준 재귀 탐색 (안정)
    # ============================================================
    root = inbox.Parent  # 보통 같은 mailbox 루트

    archive_folder = find_folder_recursive(
        root,
        ["\\archive", "online archive", "\\보관", "\\아카이브", "in-place archive"]
    )

    print(f"📅 검색 기간: {START_DATE.date()} ~ {END_DATE.date()}")
    print(f"💾 저장 위치: {SAVE_FOLDER}")

    total = 0

    # Inbox 처리
    total += process_folder(inbox, "INBOX")

    # Archive 처리
    if archive_folder:
        print("\n✅ Archive 폴더 찾음!")
        print(f"   - Name: {archive_folder.Name}")
        try:
            print(f"   - Path: {archive_folder.FolderPath}")
        except Exception:
            pass

        total += process_folder(archive_folder, "ARCHIVE")
    else:
        print("\n⚠️ [주의] Archive 폴더를 재귀 탐색으로도 찾지 못했습니다.")
        print("   -> Outlook에서 Archive가 'Online Archive - 이름'처럼 별도 사서함(Store)인지 확인 필요")
        print("   -> 그 경우 'Store 전체 순회' 버전으로 100% 잡을 수 있음")

    print(f"\n🎉 전체 완료! 총 {total}개의 파일을 다운로드했습니다.")


if __name__ == "__main__":
    download_inbox_and_archive()
