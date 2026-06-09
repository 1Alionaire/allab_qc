import json
from pathlib import Path
import pandas as pd

analysts = ['OV', 'KK', 'AB', 'VC']

# grouped_date_file_path = Path(str(Path.cwd()) + '/analysts_grouped.json')
# output_file_path = Path(str(Path.cwd()) + '/all_analyst_dup_data.json')

# source_path = Path(r"C:\Python\allab\allab_qc\data\qc_raw_data.json")
# output_dir  = Path(r"C:\Python\allab\allab_qc\data\monthly")
# output_dir.mkdir(parents=True, exist_ok=True)

source_path = Path(str(Path.cwd()) +  "/qc_raw_data.json")
output_dir  = Path(str(Path.cwd()) +  "/monthly")
output_dir.mkdir(parents=True, exist_ok=True)

df = pd.read_json(source_path, orient="index", encoding="utf-8")
df.index.name = "project_number"
df = df.reset_index()

# Парсим дату и добавляем колонку месяца в формате "2026-01"
df["date"]  = pd.to_datetime(df["date"], errors="coerce")
df = df.dropna(subset=["date"])                                # отбрасываем строки с битой датой
df["month"] = df["date"].dt.to_period("M").astype(str)         # "2026-01", "2026-02", ...

# Перебираем уникальные месяцы и сохраняем по файлу на каждый
for month, month_df in df.groupby("month"):
    result = {}
    for analyst in analysts:
        analyst_df = month_df[month_df["analyst"] == analyst]

        if analyst_df.empty:
            result[analyst] = []
            continue

        agg = (
            analyst_df
            .groupby("date", as_index=False)[["plm_count", "nob_count", "tem_count"]]
            .sum()
        )
        # datetime → строка, чтобы json.dump не ругался
        agg["date"] = agg["date"].dt.strftime("%Y-%m-%d")
        result[analyst] = agg.to_dict(orient="records")

    output_path = output_dir / f"analysts_{month}.json"
    with output_path.open("w", encoding="utf-8") as f:
        json.dump(result, f, indent=4, ensure_ascii=False)

    print(f"Сохранено: {output_path.name}")