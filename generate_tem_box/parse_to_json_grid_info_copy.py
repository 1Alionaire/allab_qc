import json
from pathlib import Path
from openpyxl import load_workbook

script_dir = Path(__file__).resolve().parent

with open(script_dir / "tem_clean_sorted_data.json", "r", encoding="utf-8") as f:
    input_data = json.load(f)

# индекс для мгновенного поиска; setdefault сохраняет первое совпадение (как next())
index = {}
for item in input_data:
    index.setdefault((str(item.get("project")), str(item.get("sample"))), item)

wb = load_workbook(script_dir / "grid_id_QC_boxes_TEM.xlsx",
                   data_only=True, read_only=True)
ws = wb["QC_boxes"]

rows = list(ws.values)   # ЕДИНСТВЕННЫЙ проход по файлу, дальше работаем с памятью
wb.close()

# колонки в кортеже: A=0 B=1 C=2 D=3 E=4 F=5 G=6 H=7 I=8
tem_grid_data = []

# боксы ищем по содержимому A, а не шагом 25 — не собьётся из-за строк-разделителей
box_starts = [i for i, r in enumerate(rows)
              if r and r[0] and str(r[0]).startswith("Box #")]

for start in box_starts:
    box_number = str(rows[start][0]).split("#")[1].strip()
    # if int(box_number) > 20:
    #     break
    block = rows[start:start + 25]

    # левая половина (B,C,D,E) и правая (F,G,H,I) — одна и та же логика
    for cell_col, letter_col, proj_col, samp_col in ((1, 2, 3, 4), (5, 6, 7, 8)):
        grid_id = None
        for off, r in enumerate(block):
            if r[cell_col] is not None:
                grid_id = r[cell_col]
            if r[proj_col] is None:
                continue

            grid_1 = str(r[letter_col]) + str(grid_id)
            if grid_1 == "A6":          # дубль граничного сэмпла, уже учтён как E5
                continue

            item = index.get((str(r[proj_col]), str(r[samp_col])))
            if item is None:
                print(f"NOT FOUND: {r[proj_col]}/{r[samp_col]} (box {box_number})")
                continue

            item["Box Number"] = box_number + "rd"
            item["Grid_1"] = grid_1
            item["Grid_2"] = "A6" if grid_1 == "E5" else \
                str(block[off + 1][letter_col]) + str(grid_id)
            tem_grid_data.append(item)

with open(script_dir / "tem_grid_data.json", "w", encoding="utf-8") as f:
    json.dump(tem_grid_data, f, indent=4, ensure_ascii=False)

print(f"Готово: {len(tem_grid_data)} записей")