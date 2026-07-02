import sys
import json
from collections import Counter

def lab_id_sort_key(lab_id):
    # lab_id = str(item.get("lab id", ""))
    lab_id = lab_id.replace(' ', '')
    try:
        main_part, number_part, sample_part = lab_id.split("-")
        return int(main_part), int(number_part), int(sample_part)
    except ValueError:
        return 999999999, 999999999

# =============================================== filter blanks
# input_data = None
# input_json_path = "tem_total_data.json"
# with open(input_json_path, "r", encoding="utf-8") as f:
#     input_data = json.load(f)

# tem_without_blanks = [element for element in input_data if element['sample'] != 'bl']
# output_json_path = "tem_without_blanks.json"
# with open(output_json_path, "a", encoding="utf-8") as f:
#     f.write(json.dumps(tem_without_blanks, indent=4, ensure_ascii=False))

# ============================================== filter duplicates
# input_data = None
# input_json_path = "tem_without_blanks.json"
# with open(input_json_path, "r", encoding="utf-8") as f:
#     input_data = json.load(f)

# lab_id_counts = Counter(
#     item.get("lab id")
#     for item in input_data
# )

# filtered_data = [
#     item
#     for item in input_data
#     if lab_id_counts[item.get("lab id")] == 1
# ]

# output_json_path = "tem_without_duplicates.json"
# with open(output_json_path, "a", encoding="utf-8") as f:
#     f.write(json.dumps(filtered_data, indent=4, ensure_ascii=False))

# =============================================== Sorting Projects through LAB ID
# input_data = None
# input_json_path = "tem_without_duplicates.json"
# with open(input_json_path, "r", encoding="utf-8") as f:
#     input_data = json.load(f)

# sorted_data = sorted(
#     input_data,
#     key=lambda item: lab_id_sort_key(item.get("lab id", ""))
# )

# output_json_path = "tem_sorted.json"
# with open(output_json_path, "a", encoding="utf-8") as f:
#     f.write(json.dumps(sorted_data, indent=4, ensure_ascii=False))

# =============================================== Sorting Projects through Date
# input_data = None
# input_json_path = "tem_without_duplicates.json"
# with open(input_json_path, "r", encoding="utf-8") as f:
#     input_data = json.load(f)

# sorted_data = sorted(
#     input_data,
#     key=lambda item: item.get("date", "")
# )

# output_json_path = "tem_sorted.json"
# with open(output_json_path, "a", encoding="utf-8") as f:
#     f.write(json.dumps(sorted_data, indent=4, ensure_ascii=False))