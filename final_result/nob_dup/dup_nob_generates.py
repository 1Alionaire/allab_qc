import json
from pathlib import Path
import math
import random
import copy
import os
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))
from utils.utilities import lab_id_sort_key, random_calc_point_asb

from utils.utilities import lab_id_sort_key, random_calc_point_asb
      
def get_plm_nob_random_sample_from_project(samples_list, analyst): 
    samples_list_duplicate = copy.deepcopy(samples_list)
    
    while True:
        random_index = random.randint(0, len(samples_list_duplicate) - 1)
        removed_element = samples_list_duplicate.pop(random_index)
        if removed_element['Type Asb 1 Option'] != 'PS':
            if removed_element['Method'] != 'None':
                original_sample = removed_element
                break

    duplicate_sample = copy.deepcopy(original_sample)        #random_calc_point_asb(samples_list[i][f'Point {j}'])

    duplicate_sample['Client ID'] = 'D' + original_sample['Client ID'] + analyst

    if duplicate_sample['Type Asb 1 Option'] != 'NAD' and duplicate_sample['Type Asb 1 Option'] != 'None':
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

        duplicate_sample['Percent 1 Option'] = round(((intermediate_percent_result * float(duplicate_sample['Total Residue'])) / 100), 2) #change

    if duplicate_sample['Type Asb 1 Option'] == 'None':
        duplicate_sample['Type Asb 1 Option'] = 'NAD'
        duplicate_sample['Percent 1 Option'] = 'NAD'
        
        original_sample['Type Asb 1 Option'] = 'NAD'
        original_sample['Percent 1 Option'] = 'NAD'

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
    monthly_pivot_dir = script_dir.parent / "monthly"
    resource_dir = script_dir.parent / "resource"

    path = Path(resource_dir / 'qc_raw_data.json')

    with path.open("r", encoding="utf-8") as f:
        raw_data = json.load(f) 

    for file in monthly_pivot_dir.glob("*.json"):
        file_base_name = os.path.basename(file)

        with file.open("r", encoding="utf-8") as f:
            grouped_data = json.load(f)

        data_per_month = { 'total':[]}
        for analyst, info in grouped_data.items():

            data_per_analyst = []
            for record in info:
                nob_total_samples = 0       # total count samples per day. 
                # filter all projects by day
                projects_by_date = {key: value for key, value in raw_data.items() if (value['date'][:-8].strip() == record['date']
                                                                                      and value['nob_count'] > 0) 
                                                                                      and value['analyst'] == analyst }
                nob_counter  = 0         # this counter need only for first blank sample in first project of the day
                #iterate through all projects in this day

                for project, project_info in projects_by_date.items():

                    if project_info['nob_count'] > 0 and project_info['nob_analysis']:
                        before = nob_total_samples
                        after = before + project_info['nob_count']

                        duplicates_before = before // 50
                        duplicates_after = after // 50
                        duplicates_to_add = duplicates_after - duplicates_before

                        if nob_counter == 0:
                            duplicates_to_add += 1

                        if duplicates_to_add > 0:
                            for i in range(duplicates_to_add):
                                sample_duplicate = get_plm_nob_random_sample_from_project(project_info['nob_analysis'], analyst)

                                duplicate_record = {
                                    'date': project_info['date'], 
                                    'project': project,
                                    'lab id': sample_duplicate['whole_original']['Lab ID'],
                                    'file_name': project_info['file_name'],
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

                                for l in project_info['plm_analysis']:
                                    if l['Client ID'] == duplicate_record['sample']:
                                        project_info['plm_analysis'].remove(l)

                        nob_total_samples = after
                        nob_counter+=1

            data_per_analyst_sorted = sorted(data_per_analyst, key=lab_id_sort_key)
            data_per_month[analyst] = data_per_analyst_sorted

        unsorted_total_data = data_per_month['total']
        sorted_total_data = sorted(unsorted_total_data, key=lab_id_sort_key)

        new_data_per_month = {key: value for key, value in data_per_month.items() if key != 'total'}
        new_data_per_month['total'] = sorted_total_data

        output_file_path = script_dir /  file_base_name

        with open(output_file_path, "a", encoding="utf-8") as file:
            json.dump(new_data_per_month, file, indent=2, ensure_ascii=False)

if __name__ == "__main__":
    main_start()