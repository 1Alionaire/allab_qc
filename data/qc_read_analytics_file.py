from pathlib import Path
import pandas as pd

source_path = Path(r"C:\Python\allab\allab_qc\data\KK_data.jsonl")

df = pd.read_json(source_path, encoding="utf-8")

df[["project_date_part", "project_seq"]] = (
    df["project_number"]
    .str.split("-", expand=True)
    .astype(int)
)

df_sorted = df.sort_values(
    by=["project_date_part", "project_seq"],
    ascending=[True, True]
)

