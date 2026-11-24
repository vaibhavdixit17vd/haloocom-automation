import os
import time
from pathlib import Path
import pandas as pd

# Define paths
base_dir = Path(r"C:\Users\vaibhav.dixit\company-revenue-mapper\data")
excel_path = base_dir / "master_revenue_dataset.xlsx"
csv_path = base_dir / "master_revenue_dataset.csv"

if not excel_path.exists():
    # Fallback if master file is in the root directory
    excel_path = Path("master_revenue_dataset.xlsx")
    csv_path = Path("master_revenue_dataset.csv")

print(f"Opening Excel file: {excel_path}...")
start_time = time.time()

# 1. Inspect all sheet names in the workbook
xls = pd.ExcelFile(excel_path)
print(f"Found {len(xls.sheet_names)} sheet(s): {xls.sheet_names}")

all_sheets = []
total_rows = 0

# 2. Iterate through each sheet and combine
for sheet_name in xls.sheet_names:
    print(f"Reading sheet: '{sheet_name}'...")
    df_sheet = pd.read_excel(xls, sheet_name=sheet_name)

    # Strip whitespace from column headers
    df_sheet.columns = [str(col).strip() for col in df_sheet.columns]

    # Skip empty sheets
    if df_sheet.empty:
        print(f"  -> Sheet '{sheet_name}' is empty. Skipping.")
        continue

    print(f"  -> Loaded {len(df_sheet):,} rows and {len(df_sheet.columns)} columns.")
    total_rows += len(df_sheet)
    all_sheets.append(df_sheet)

if not all_sheets:
    raise ValueError("No readable data found in any sheet!")

# 3. Concatenate all sheets into one unified DataFrame
print("\nCombining all sheets into a single master dataset...")
df_consolidated = pd.concat(all_sheets, ignore_index=True)

# 4. Save to CSV
print(f"Saving {len(df_consolidated):,} total rows to CSV at: {csv_path}...")
df_consolidated.to_csv(csv_path, index=False, encoding="utf-8-sig")

elapsed_time = time.time() - start_time
print("=" * 60)
print(f"✅ SUCCESS! Consolidated CSV created in {elapsed_time:.2f} seconds.")
print(f"Output File: {csv_path.resolve()}")
print(f"Total Combined Rows: {len(df_consolidated):,}")
print("=" * 60)