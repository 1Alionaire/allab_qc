import json
from pathlib import Path
import pandas as pd

path = Path(r"C:\Python\allab\allab_qc\data\qc_data_cleaned.jsonl")

df = pd.read_json(path, orient="index", encoding="utf-8")
df.index.name = "project_number"      # называем индекс
df = df.reset_index()                 # делаем номер проекта обычной колонкой

# print(df.head())
# print(df.shape)
# print(df.dtypes)

analyst_df = df[df['analyst'] == 'No Analyst']
# empty_df = df[df['project_number']].isna()

print(analyst_df)
# print(empty_df)
# analyst_array = df['analyst'].unique()
# ['OV' 'No Analyst' 'KK' 'AB' 'VC']
# print(analyst_array)



# output_file = Path(self.folder_path) / "qc_data.jsonl"



# print(analyst_df.shape)
# print(analyst_df.dtypes)
# print(analyst_df['analyst'].head())
# python C:\Python\allab\allab_qc\qc_open_json_file.py

