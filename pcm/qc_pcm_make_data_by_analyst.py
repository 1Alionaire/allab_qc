import json
from pathlib import Path
import pandas as pd

BASE_DIR = Path(__file__).resolve().parent

source_path = BASE_DIR / 'qc_pcm_wrong_files.json'

with source_path.open('r', encoding='utf-8') as f:
    data = json.load(f)

df = pd.DataFrame.from_dict(data, orient='index')
df.index.name = 'sample_id'
df = df.reset_index()    

print(df['analyst'].unique())

# for analyst in df['analyst'].unique():
#     analyst_df = df[df["analyst"] == analyst]

#     print(analyst_df.head())
# for analyst in analysts:
#     analyst_df = df[df["analyst"] == analyst]

#     analyst_df = (
#         analyst_df
#         .groupby("date", as_index=False)[["plm_count", "nob_count", "tem_count"]]
#         .sum()
#     )

#     result[analyst] = analyst_df.to_dict(orient="records")

# with open(output_path, "w", encoding="utf-8") as file:
#     json.dump(result, file, indent=4, ensure_ascii=False)