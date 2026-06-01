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
      
def get_plm_nob_random_sample_from_project(samples_list):

    asb_types = ['Chrysotile', 'Amosite', 'Anthophillite', 
                 'Actinolite', 'Crocidolite', 'Tremolite']
    
    samples_list_duplicate = copy.deepcopy(samples_list)
    
    while True:
        random_index = random.randint(0, len(samples_list_duplicate) - 1)
        removed_element = samples_list_duplicate.pop(random_index)
        if removed_element['Type Asb 1 Option'] != 'PS':
            original_sample = removed_element
            break

    duplicate_sample = copy.deepcopy(original_sample)        #random_calc_point_asb(samples_list[i][f'Point {j}'])

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
                if duplicate_sample[f'Point {j}'] == 'None' or duplicate_sample[f'Point {j}'] == '50':
                    pass
                else:
                    add_asb_glass += 1
                    sum_points += int(duplicate_sample[f'Point {j}'])
        
            intermediate_percent_result = round((((4 + add_asb_glass) / sum_points) * 100), 2)

        duplicate_sample['Percent 1 Option'] = intermediate_percent_result


    return {
        'sample': original_sample['Client ID'],
        '1st_analyst_asb_type': original_sample['Type Asb 1 Option'], 
        '1st_analyst_asb_percent': original_sample['Percent 1 Option'], 
        'dup_asb_type': duplicate_sample['Type Asb 1 Option'], 
        'dup_asb_percent': duplicate_sample['Percent 1 Option'],
        'whole_original':  original_sample,
        'whole_duplicate': duplicate_sample
    }

def main_start():
    script_dir = Path(__file__).resolve().parent
    monthly_dir = script_dir.parent / "monthly"
    path = Path(str(Path.cwd()) + '/qc_raw_data.json')
    total_info = {}

    with path.open("r", encoding="utf-8") as f:
        raw_data = json.load(f) 

    folder = Path(str(Path.cwd()) + '/monthly')

    for file in folder.glob("analysts_2026-01.json"):
    # for file in folder.glob("*.json"):
        print(f'file: {file}')

        with file.open("r", encoding="utf-8") as f:
            grouped_data = json.load(f)

        data_per_month = { 'total':[]}
        for analyst, info in grouped_data.items():

            data_per_analyst = []
            for record in info:
                limit = 100     # how many blank samples per day accroding audit
                plm_total_samples = 0       # total count samples per day. 
                # filter all projects by day
                projects_by_date = {key: value for key, value in raw_data.items() if (value['date'][:-8].strip() == record['date'] and value['plm_count'] > 0) }
                plm_counter  = 0         # this counter need only for first blank sample in first project of the day
                #iterate through all projects in this day

                for project, project_info in projects_by_date.items():
                    samples_in_project = []

                    if project_info['plm_count'] > 0 and project_info['plm_analysis']:
                        before = plm_total_samples
                        after = before + project_info['plm_count']

                        duplicates_before = before // 50
                        duplicates_after = after // 50
                        duplicates_to_add = duplicates_after - duplicates_before

                        if plm_counter == 0:
                            duplicates_to_add += 1

                        if duplicates_to_add > 0:
                            sample_duplicate = get_plm_nob_random_sample_from_project(project_info['plm_analysis'])

                            duplicate_record = {
                                'date': project_info['date'], 
                                'project': project,
                                'sample': sample_duplicate['sample'],
                                '1st_analyst_name': project_info['analyst'],
                                '1st_analyst_asb_type': sample_duplicate['1st_analyst_asb_type'],
                                '1st_analyst_asb_percent': sample_duplicate['1st_analyst_asb_percent'],
                                'dup_name': project_info['analyst'], 
                                'dup_asb_type': sample_duplicate['dup_asb_type'],
                                'dup_asb_percent': sample_duplicate['dup_asb_percent'],
                                'whole_original': sample_duplicate['whole_original'],
                                'whole_duplicate': sample_duplicate['whole_duplicate']
                            }
                            data_per_analyst.append(duplicate_record)
                            data_per_month['total'].append(duplicate_record)

                        plm_total_samples = after
                        plm_counter+=1

            data_per_month[analyst] = data_per_analyst

        output_file_path = str(file) + '1.json'

        with open(output_file_path, "a", encoding="utf-8") as file:
                json.dump(data_per_month, file, indent=2, ensure_ascii=False)

if __name__ == "__main__":
    main_start()