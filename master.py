import pandas as pd
import glob
import os
import warnings

# 1. 스타일 관련 경고 메시지 차단 (깔끔한 터미널을 위해)
warnings.filterwarnings("ignore", message="Workbook contains no default style*")

def process_reports(folder_path):
    # 경로 설정
    # 합쳐질 마스터 파일 이름
    master_file_name = "Master_Applications_Archive.xlsx" 
    # 정보를 채워넣어야 할 타겟 파일 (예: LITNUM 결과표 등)
    assessment_file_name = "assessment_data.xlsx"
    
    master_output_path = os.path.join(folder_path, master_file_name)
    assessment_path = os.path.join(folder_path, assessment_file_name)

    print(f"📂 작업 폴더: {folder_path}")

    # --- STEP 1. 마스터 아카이브 만들기 (Active-Applications 파일 병합) ---
    search_pattern = os.path.join(folder_path, "Active-Applications*.xlsx")
    app_files = glob.glob(search_pattern)
    
    if not app_files:
        print("❌ 오류: 'Active-Applications'로 시작하는 엑셀 파일을 찾지 못했습니다.")
        return
    
    print(f"ℹ️ 총 {len(app_files)}개의 리포트 파일을 병합 중입니다...")

    app_list = []
    for f in app_files:
        try:
            # 3번째 줄 헤더(header=2) 적용
            temp_df = pd.read_excel(f, header=2)
            temp_df.columns = temp_df.columns.str.strip()
            app_list.append(temp_df)
        except Exception as e:
            print(f"⚠️ {os.path.basename(f)} 읽기 실패: {e}")
    
    if not app_list:
        print("❌ 병합할 데이터가 없습니다.")
        return

    # 모든 파일 합치기 + 중복 제거 (지원번호 기준)
    master_archive = pd.concat(app_list, ignore_index=True).drop_duplicates(subset=['Adm Appl Nbr'], keep='last')
    
    # 이름 합치기 (조회를 위해 Full_Name 생성)
    master_archive['Full_Name'] = (master_archive['Prefered First Name'].fillna('') + " " + 
                                   master_archive['Prefered Last Name'].fillna('')).str.lower().str.strip()

    # 마스터 파일 우선 저장
    master_archive.to_excel(master_output_path, index=False)
    print(f"✅ [1단계 완료] 마스터 아카이브 생성됨: {master_file_name} (총 {len(master_archive)}건)")

    # --- STEP 2. 평가 데이터(Assessment Data) 복구 ---
    # 파일이 있을 때만 실행합니다!
    if os.path.exists(assessment_path):
        print(f"🔍 '{assessment_file_name}' 파일을 발견했습니다. 데이터 복구를 시작합니다...")
        try:
            assessment_df = pd.read_excel(assessment_path)
            assessment_df['Name_Lower'] = assessment_df['Name'].str.lower().str.strip()
            
            # 매핑용 딕셔너리 생성
            id_map = master_archive.set_index('Nsn Student Number')['Emplid'].to_dict()
            name_to_id_map = master_archive.set_index('Full_Name')['Emplid'].to_dict()
            prog_map = master_archive.set_index('Emplid')['Acad Prog'].to_dict()

            # 빈 sis_id 찾기
            mask = (assessment_df['sis_id'] == 'Not Found') | (assessment_df['sis_id'].isna())
            
            # 복구 시도 (NSN -> 이름 순서)
            assessment_df.loc[mask, 'sis_id'] = assessment_df.loc[mask, 'NSN'].map(id_map)
            mask_still_missing = (assessment_df['sis_id'].isna()) | (assessment_df['sis_id'] == 'Not Found')
            assessment_df.loc[mask_still_missing, 'sis_id'] = assessment_df.loc[mask_still_missing, 'Name_Lower'].map(name_to_id_map)
            
            # 프로그램 정보 업데이트
            assessment_df['Programme'] = assessment_df['sis_id'].map(prog_map).fillna(assessment_df['Programme'])

            # 결과 저장
            result_path = os.path.join(folder_path, "Fixed_Assessment_Results.xlsx")
            assessment_df.drop(columns=['Name_Lower']).to_excel(result_path, index=False)
            print(f"✅ [2단계 완료] 수정된 결과표 저장됨: Fixed_Assessment_Results.xlsx")
            
        except Exception as e:
            print(f"⚠️ 평가 데이터 처리 중 에러 발생: {e}")
    else:
        print(f"ℹ️ 참고: '{assessment_file_name}' 파일이 폴더에 없어 2단계(매핑)는 건너뜁니다.")

    print("-" * 30)
    print(f"🎉 모든 작업이 끝났습니다!")

# ==========================================
# 실행부
# ==========================================
if __name__ == "__main__":
    # 조나단님의 폴더 경로
    my_folder = r"C:\Users\wooin\Documents\student-applications\files"
    process_reports(my_folder)