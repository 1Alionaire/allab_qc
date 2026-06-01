import json
from pathlib import Path
import pandas as pd

analysts = ['OV', 'KK', 'AB', 'VC']
 
source_path = Path(r"C:\Python\allab\allab_qc\data\qc_raw_data.json")

df = pd.read_json(source_path, orient="index", encoding="utf-8")
df.index.name = "project_number"      # называем индекс
df = df.reset_index()                 # делаем номер проекта обычной колонкой

output_path = Path(r"C:\Python\allab\allab_qc\data\analysts_grouped.json")

result_list = []
result = {}

for analyst in analysts:
    analyst_df = df[df["analyst"] == analyst]

    analyst_df = (
        analyst_df
        .groupby("date", as_index=False)[["plm_count", "nob_count", "tem_count"]]
        .sum()
    )

    result[analyst] = analyst_df.to_dict(orient="records")

with open(output_path, "w", encoding="utf-8") as file:
    json.dump(result, file, indent=4, ensure_ascii=False)