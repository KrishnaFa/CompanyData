"""Quick verification: count prediction levels used and show sample rows."""
import sys
import os
# Insert project root (parent of scratch/) into path so analyzer can be imported
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pandas as pd
import re
from analyzer import analyze_file_path, resolve_input_xlsx

input_path = resolve_input_xlsx(None)
result = analyze_file_path(str(input_path))
df = result["processed_df"]

# Count prediction levels for imputed rows
imputed = df[df["R1 Source"] == "Imputed"].copy()
imputed["Level"] = imputed["R1 Imputation Method"].apply(
    lambda x: re.search(r"Predicted (L\d)", str(x)).group(1) if re.search(r"Predicted (L\d)", str(x)) else "Other"
)

print("=" * 60)
print(f"Total candidates: {len(df)}")
print(f"R1 dates: Original/Yr-Corrected = {len(df[df['R1 Source'].isin(['Original','Yr Corrected'])])}")
print(f"R1 dates: Imputed = {len(imputed)}")
print(f"R1 dates: Missing (no impute needed) = {df['R1 Date (Final)'].isna().sum()}")
print()
print("Prediction levels used:")
print(imputed["Level"].value_counts().to_string())

print()
print("Sample imputed rows (5):")
sample = imputed.sample(min(5, len(imputed)), random_state=42)[
    ["Name", "Current Designation", "Skills", "R0", "R1 Date (Final)", "R1 Imputation Method"]
]
for _, row in sample.iterrows():
    print(f"  {row['Name']} | {row['Current Designation']} | {row['Skills']}")
    print(f"    R0={row['R0'].date() if pd.notna(row['R0']) else 'N/A'}  R1={row['R1 Date (Final)'].date() if pd.notna(row['R1 Date (Final)']) else 'N/A'}")
    print(f"    Method: {str(row['R1 Imputation Method'])[:100]}")
    print()
