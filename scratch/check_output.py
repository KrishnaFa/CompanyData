import pandas as pd
import re

df = pd.read_excel("Supply_Mapping_PowerBI_Ready.xlsx", sheet_name="Recruitment Timeline", header=1)
print("Columns:", df.columns.tolist())
