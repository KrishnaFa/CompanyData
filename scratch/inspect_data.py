import pandas as pd
df = pd.read_excel("Supply Mapping.xlsx")
print("Columns:")
print(df.columns.tolist())
print("\nFirst 3 rows:")
print(df.head(3).to_dict(orient="records"))
print("\nUnique designations:")
print(df["Current Designation"].value_counts().head(20))
