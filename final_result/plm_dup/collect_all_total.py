from pathlib import Path
import json
import logging


logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s'
)

folder = Path("data")
name_output_file = 'PLM_DUP.json'
all_total = []

script_dir = Path(__file__).resolve()

excel_result_dir = script_dir.parent / "backup"
output_json_file = excel_result_dir / name_output_file

for file in excel_result_dir.glob("*.json"):
    if name_output_file in str(file):
        continue
    with file.open("r", encoding="utf-8") as f:
        data = json.load(f)

    all_total.extend(data.get("total", []))
logging.info(all_total)
with open(output_json_file, "a", encoding="utf-8") as file:
    json.dump(all_total, file, indent=2, ensure_ascii=False)


