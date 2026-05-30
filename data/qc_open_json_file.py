import json
from pathlib import Path
import math
import random

def random_calc_point_asb(inp_str):
    if inp_str == '50':
        return '50'
    operation = random.choice(['-', '+'])
    if operation == '-':
        value = random.choice([1, 2, 3])
        return str(int(inp_str) - value)
    else:
        value = random.choice([1, 2, 3])
        if (int(inp_str) + value) < 50:
            return str(int(inp_str) + value)
        else:
            return '49'

analysts = ['OV', 'KK', 'AB', 'VC']
def get_plm_nob_random_sample_from_project(samples_list, type, analyst):
    other_analysts = analysts.remove(analyst)
    rep_analyst = random.choice(other_analysts)

    original_sample = random.choice(samples_list)
    duplicate_sample = original_sample
    if type == 'dup':
        duplicate_sample['Client ID'] = 'D' + original_sample['Client ID'] + rep_analyst
        if original_sample['Type 1'] == 'PS':
            original_index = samples_list.index(original_sample)
            for i in range((original_index - 1), -1, -1):
                if samples_list[i]['Type 1'] != 'PS':
                    for j in range(1, 9):
                        if duplicate_sample[f'Type {j}'] != 'None':
                            duplicate_sample[f'Type {j}'] = samples_list[i][f'Type {j}']
                            duplicate_sample[f'Point {j}'] = random_calc_point_asb(samples_list[i][f'Point {j}'])
                   
        

        return duplicate_sample
    else:
        duplicate_sample['Client ID'] = 'R' + original_sample['Client ID'] + rep_analyst
        return duplicate_sample

analyst = 'VC'
types_analysis = ["plm_count", "nob_count","tem_count",]


path = Path(fr"C:\Python\allab\allab_qc\data\{analyst}_data.jsonl")
output_file = Path(fr"C:\Python\allab\allab_qc\data\{analyst}_dup_data.jsonl")

grouped_date_file = Path(r"C:\Python\allab\allab_qc\data\analysts_grouped.json")

final_result = {}

with path.open("r", encoding="utf-8") as f:
    data = json.load(f)       

with grouped_date_file.open("r", encoding="utf-8") as f:
    grouped_data_analysts = json.load(f) 

grouped_data_analysts = grouped_data_analysts[analyst]


for record in grouped_data_analysts:
    # how many blank samples per day accroding audit
    limit = 100
    # total count samples per day. 
    nob_total_samples = 0
    # filter all projects by day
    projects_by_date = {key: value for key, value in data.items() if value['date'] == record['date'] }
    # this counter need only for first blank sample in first project of the day
    counter = 0


    #iterate through all projects in this day
    for project, project_info in projects_by_date.items():
        final_project_info = {'plm_rep_samples': [], 'plm_dup_samples': [], 
                        'nob_rep_samples': [], 'nob_dup_samples': [],
                        'tem_rep_samples': [], 'tem_dup_samples': [], 
                        'file_name':project_info['file_name'], 'analyst':project_info['analyst']}

        if project_info['plm_count'] > 0:
            pass

        if project_info['nob_count'] > 0:
            # general
            before = nob_total_samples
            after = before + project_info['nob_count']

            # counting blanks
            blanks_before = before // 100
            blanks_after = after // 100
            blanks_to_add = blanks_after - blanks_before

            # counting replicates
            replicates_before = before // 15
            replicates_after = after // 15
            replicates_to_add = replicates_after - replicates_before

            duplicates_before = before // 50
            duplicates_after = after // 50
            duplicates_to_add = duplicates_after - duplicates_before

            if counter == 0:
                blanks_to_add +=1
                replicates_to_add += 1
                duplicates_to_add += 1

            print('=' * 25)
            print(project_info['date'])
            print(f'before: {before}')
            print(f'after: {after}')
            print('-' * 25)
            print(f'blanks_before: {blanks_before}')
            print(f'blanks_after: {blanks_after}')
            print(f'blanks_to_add: {blanks_to_add}')
            print('-' * 25)
            print(f'replicates_before: {replicates_before}')
            print(f'replicates_after: {replicates_after}')
            print(f'replicates_to_add: {replicates_to_add}')
            print('-' * 25)
            print(f'duplicates_before: {duplicates_before}')
            print(f'duplicates_after: {duplicates_after}')
            print(f'duplicates_to_add: {duplicates_to_add}')
            

            if blanks_to_add > 0:
                for i in range(blanks_to_add):
                    final_project_info['nob_rep_samples'].append({
                            "Client ID": "bl",
                            "Lab ID": "1",
                            "Layer": "",
                            "Color": "",
                            "Texture": "Non-Fibs",
                            "Homogeneity": "",
                            "Morphology": "",
                            "RI II Type 1": "",
                            "RI II Type 2": "",
                            "RI ┴ Type 1": "",
                            "RI ┴ Type 2": "",
                            "Sign of \nElongation Type 1": "",
                            "Sign of \nElongation Type 2": "",
                            "Extinction \nAngle Type 1": "",
                            "Extinction \nAngle Type 2": "",
                            "Pleochroism /\nColor Type 1": "",
                            "Pleochroism /\nColor Type 2": "",
                            "Birefringence Type 1": "",
                            "Birefringence Type 2": "",
                            "Other Fibers": "",
                            "Property": "",
                            "% Non-\nAsbestos": "",
                            "Type 1": "NAD",
                            "Point 1": "50",
                            "Type 2": "NAD",
                            "Point 2": "50",
                            "Type 3": "NAD",
                            "Point 3": "50",
                            "Type 4": "NAD",
                            "Point 4": "50",
                            "Type 5": "",
                            "Point 5": "",
                            "Type 6": "",
                            "Point 6 ": "",
                            "Type 7": "",
                            "Point 7": "",
                            "Type 8": "",
                            "Point 8": "",
                            "Percent For type 1": "NAD",
                            "Asb Type For type 1": "NAD",
                            "Percent For type 2": "",
                            "Asb Type For type 2": "",
                            "Vermiculite": "ND",
                            "Method": "198.6",
                            "Undesolved Materials": "",
                            "Total Residue": ""
                        })
                    
            if replicates_to_add > 0:
                for i in range(replicates_to_add):
                    final_project_info['nob_rep_samples'].append(get_plm_nob_random_sample_from_project(project_info['nob_analysis'], 'rep', project_info['analyst']))

            if duplicates_to_add > 0:
                for i in range(duplicates_to_add):
                    final_project_info['nob_dup_samples'].append(get_plm_nob_random_sample_from_project(project_info['nob_analysis'], 'dup', project_info['analyst']))

            final_project_info['blank_nob_count'] = blanks_to_add
            final_project_info['rep_nob_count'] = replicates_to_add
            final_project_info['dup_nob_count'] = duplicates_to_add
            nob_total_samples = after
            counter+=1
        final_result[project] = final_project_info


# print(final_result)

with open(output_file, "w", encoding="utf-8") as file:
    json.dump(final_result, file, indent=4, ensure_ascii=False)