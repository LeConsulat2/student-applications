import pandas as pd
import warnings
import sys
from pathlib import Path

warnings.filterwarnings("ignore")

# ================= CONFIG =================
USER_DOCS = Path.home() / "Documents"
OUTPUT_FOLDER = USER_DOCS / "student-applications"
RAW_FILES_FOLDER = OUTPUT_FOLDER / "files"

# Set your exact desired output filename here
MASTER_OUTPUT = OUTPUT_FOLDER / "Active Application Details.xlsx"

APPLICATION_COL = "Adm Appl Nbr"
# ========================================

def load_and_clean_excel(file_path: Path) -> pd.DataFrame:
    """Load one Excel file and apply the cleaning logic."""
    df = pd.read_excel(file_path, header=2)

    # Clean headers
    df.columns = df.columns.str.strip()

    # Drop ghost/unnamed columns
    df = df.loc[:, ~df.columns.str.contains(r"^Unnamed", na=False)]

    # Drop fully empty rows
    df = df.dropna(how="all", axis=0)

    # Remove disclaimer/junk rows using boolean indexing
    if not df.empty:
        df = df[~df.iloc[:, 0].astype(str).str.contains("This page provides", na=False)]

    return df

def stack_all_files() -> pd.DataFrame:
    """Finds all target Excel files, cleans them, and stacks them together."""
    # The pattern *Active*Application*.xlsx catches:
    # "Active Applications.xlsx", "Active-Applications.xlsx", "Active_Application_001.xlsx", etc.
    files = list(RAW_FILES_FOLDER.glob("*Active*Application*.xlsx"))

    if not files:
        print(f"⚠️ No files found in {RAW_FILES_FOLDER} matching 'Active Applications'.")
        sys.exit(1)

    all_dfs = []
    for file in files:
        try:
            df = load_and_clean_excel(file)
            if not df.empty:
                df["source_file"] = file.name
                all_dfs.append(df)
                print(f"✅ Loaded: {file.name}")
        except Exception as e:
            print(f"❌ Skipped {file.name}: {e}")

    if not all_dfs:
        print("⚠️ No valid data could be extracted from the files.")
        sys.exit(1)

    return pd.concat(all_dfs, ignore_index=True)

def deduplicate_by_application_number(df: pd.DataFrame) -> pd.DataFrame:
    """Removes empty application numbers and deduplicates."""
    if APPLICATION_COL not in df.columns:
        print(f"❌ Critical Error: Column '{APPLICATION_COL}' not found. Check your source files.")
        sys.exit(1)

    # Drop rows where the application number is missing, then deduplicate
    df_clean = df.dropna(subset=[APPLICATION_COL])
    return df_clean.drop_duplicates(subset=[APPLICATION_COL], keep="last")

def main():
    print("🚀 Starting Data Merge & Deduplication...\n")
    
    # Create output folder if it doesn't exist
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

    # 1. Stack everything into memory
    stacked_df = stack_all_files()
    
    # 2. Deduplicate directly (No intermediate saves)
    final_df = deduplicate_by_application_number(stacked_df)
    
    # 3. Save final output to your exact requested filename
    final_df.to_excel(MASTER_OUTPUT, index=False)
    
    print(f"\n💾 Saved final master file: {MASTER_OUTPUT.name}")
    print(f"📊 Total rows combined: {len(stacked_df)}")
    print(f"🧹 Duplicates removed: {len(stacked_df) - len(final_df)}")
    print(f"✅ Final clean rows: {len(final_df)}")
    print("\n✨ Process Complete!")

if __name__ == "__main__":
    main()