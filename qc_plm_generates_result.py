import json
from pathlib import Path
import math
import random
import copy

def random_calc_point_asb(inp_str):
    if inp_str == 'None':
        return 'None'
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

def replicate_analyst(entry_analyst):
    if entry_analyst == "AB":
        return 'KK'
    elif entry_analyst == "KK":
        return 'AB'
    elif entry_analyst == "OV":
        return 'AB'
    elif entry_analyst == "VC":
        return 'AB'
        
def get_tem_random_sample_from_project(samples_list, type, analyst):
    rep_analyst = replicate_analyst(analyst)
    original_sample = random.choice(samples_list)
    duplicate_sample = copy.deepcopy(original_sample)
    
    if type == 'dup':
        duplicate_sample['Client ID'] = 'D' + original_sample['Client ID'] + rep_analyst
    else:
        duplicate_sample['Client ID'] = 'R' + original_sample['Client ID'] + rep_analyst

def get_plm_nob_random_sample_from_project(samples_list, type, analyst):
    rep_analyst = replicate_analyst(analyst)

    original_sample = random.choice(samples_list)
    duplicate_sample = copy.deepcopy(original_sample)

    if type == 'dup':
        duplicate_sample['Client ID'] = 'D' + original_sample['Client ID'] + analyst
    else:
        duplicate_sample['Client ID'] = 'R' + original_sample['Client ID'] + rep_analyst

    if original_sample['Type Asb 1 Option'] == 'PS':
        original_index = samples_list.index(original_sample)
        for i in range((original_index - 1), -1, -1):
            if samples_list[i]['Type 1'] != 'PS':

                duplicate_sample['Type Asb 1 Option'] = samples_list[i]['Type Asb 1 Option']
                duplicate_sample['Percent 1 Option'] = samples_list[i]['Percent 1 Option']
                duplicate_sample['Type Asb 2 Option'] = samples_list[i]['Type Asb 2 Option']
                duplicate_sample['Percent 2 Option'] = samples_list[i]['Percent 2 Option']

                for j in range(1, 9):
                    if duplicate_sample[f'Type {j}'] != 'None':
                        duplicate_sample[f'Type {j}'] = samples_list[i][f'Type {j}']
                        duplicate_sample[f'Point {j}'] = samples_list[i][f'Point {j}'] #random_calc_point_asb(samples_list[i][f'Point {j}'])

    if duplicate_sample['Type Asb 1 Option'] != 'NAD':
        for j in range(1, 9):
            if duplicate_sample[f'Type {j}'] != 'None':
                duplicate_sample[f'Point {j}'] = random_calc_point_asb(original_sample[f'Point {j}'])

        sum_points = 0
        intermediate_percent_result = 0

        for j in range(1, 5):
            try:
                int_point = int(duplicate_sample[f'Point {j}'])
            except:
                int_point = 50

            sum_points +=  int_point
        
        if sum_points < 200:
            intermediate_percent_result = round(((4 / sum_points) * 100), 2)
        else:
            add_asb_glass = 0
            for j in range(5, 9):
                print(duplicate_sample[f'Point {j}'])
                if duplicate_sample[f'Point {j}'] == 'None' or duplicate_sample[f'Point {j}'] == '50':
                    pass
                else:
                    add_asb_glass += 1
                    sum_points += int(duplicate_sample[f'Point {j}'])
        
            intermediate_percent_result = round((((4 + add_asb_glass) / sum_points) * 100), 2)

        if duplicate_sample['Method'] == '198.1':
            duplicate_sample['Percent 1 Option'] = intermediate_percent_result #Change key
        else:
            duplicate_sample['Percent 1 Option'] = round(((intermediate_percent_result * float(duplicate_sample['Total Residue'])) / 100), 2) #change

    if type == 'dup':
        return {
            'sample': original_sample['Client ID'],
            '1st_analyst_name': analyst, 
            '1st_analyst_asb_type': original_sample['Type Asb 1 Option'], 
            '1st_analyst_asb_percent': original_sample['Percent 1 Option'], 
            'dup_name': analyst, 
            'dup_asb_type': duplicate_sample['Type Asb 1 Option'], 
            'dup_asb_percent': duplicate_sample['Percent 1 Option'], 
        }
    else:
        return {
            'sample': original_sample['Client ID'],
            '1st_analyst_name': analyst, 
            '1st_analyst_asb_type': original_sample['Type Asb 1 Option'], 
            '1st_analyst_asb_percent': original_sample['Percent 1 Option'], 
            'rep_name': rep_analyst, 
            'rep_asb_type': duplicate_sample['Type Asb 1 Option'], 
            'rep_asb_percent': duplicate_sample['Percent 1 Option'], 
        }

def main_start():
    path = Path(str(Path.cwd()) + '/qc_raw_data.json')
    # path = Path(str(Path.cwd()) + '/VC_data.json')
    with path.open("r", encoding="utf-8") as f:
        raw_data = json.load(f) 

    folder = Path(str(Path.cwd()) + '/monthly')
    
    total_info = {}
    for file in folder.glob("*.json"):
        print(f'file: {file}')

        with file.open("r", encoding="utf-8") as f:
            grouped_data = json.load(f)

        data_per_month = {}
        for analyst, info in grouped_data.items():

            data_per_analyst = []
            for record in info:
                all_projects_per_day = {}
                limit = 100     # how many blank samples per day accroding audit
                plm_total_samples = 0       # total count samples per day. 
                # filter all projects by day
                projects_by_date = {key: value for key, value in raw_data.items() if (value['date'][:-8].strip() == record['date'] and value['plm_count'] > 0) }
                plm_counter  = 0         # this counter need only for first blank sample in first project of the day
                #iterate through all projects in this day

                for project, project_info in projects_by_date.items():
                    samples_in_project = []

                    if project_info['plm_count'] > 0:
                        before = plm_total_samples
                        after = before + project_info['plm_count']

                        duplicates_before = before // 50
                        duplicates_after = after // 50
                        duplicates_to_add = duplicates_after - duplicates_before

                        if plm_counter == 0:
                            duplicates_to_add += 1

                        if duplicates_to_add > 0:
                            # for i in range(duplicates_to_add):
                            samples_in_project.append(duplicates_to_add)
                            
                        plm_total_samples = after
                        plm_counter+=1

                    all_projects_per_day[project] = samples_in_project

                if all_projects_per_day:
                    data_per_analyst.append(all_projects_per_day)

            data_per_month[analyst] = data_per_analyst

        output_file_path = str(file) + '1.json'

        with open(output_file_path, "a", encoding="utf-8") as file:
                json.dump(data_per_month, file, indent=2, ensure_ascii=False)


if __name__ == "__main__":
    main_start()