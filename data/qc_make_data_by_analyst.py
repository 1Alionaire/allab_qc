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
# df[["project_date_part", "project_seq"]] = (
#     df["project_number"]
#     .str.split("-", expand=True)
#     .astype(int)
# )

# df_sorted = df.sort_values(
#     by=["project_date_part", "project_seq"],
#     ascending=[True, True]
# )

# Делаем данные по аналитикам
# for analyst in analysts:
#     filtered_df = df[df["analyst"] == analyst]
#     filtered_sorted_df = filtered_df.sort_values(
#         by=["project_date_part", "project_seq"],
#         ascending=[True, True]
#     )
#     filtered_sorted_df.to_json(f"{analyst}_data.jsonl", orient="records", indent=4, force_ascii=False)


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