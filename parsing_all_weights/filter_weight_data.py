import sys
import json
from collections import Counter

# ============================================== filter duplicates

# def remove_identical_elements(items):
#     seen = set()
#     result = []

#     for item in items:
#         key = tuple(sorted(item.items()))

#         if key not in seen:
#             seen.add(key)
#             result.append(item)

#     return result

# input_data = None
# input_json_path = "weight_data.json"
# with open(input_json_path, "r", encoding="utf-8") as f:
#     input_data = json.load(f)

# print(len(input_data))

# unique_data = remove_identical_elements(input_data)

# print(len(unique_data)) 

# output_json_path = "weight_data_no_dups.json"
# with open(output_json_path, "a", encoding="utf-8") as f:
#     f.write(json.dumps(unique_data, indent=4, ensure_ascii=False))

# ============================================== filter values where residue, caco3 and orgainc more than 0 and less than 100 
input_data = None
input_json_path = "weight_data_no_dups.json"
with open(input_json_path, "r", encoding="utf-8") as f:
    input_data = json.load(f)

print(len(input_data))

filtered_correct_data = [element for element in input_data if element['percent_organic'] < 100 and element['percent_organic'] > 0]
filtered_correct_data = [element for element in filtered_correct_data if element['percent_caco3'] < 100 and element['percent_caco3'] > 0]
filtered_correct_data = [element for element in filtered_correct_data if element['percent_residue'] < 100 and element['percent_residue'] > 0]

print(len(filtered_correct_data)) 

output_json_path = "correct_weight.json"
with open(output_json_path, "a", encoding="utf-8") as f:
    f.write(json.dumps(filtered_correct_data, indent=4, ensure_ascii=False))


