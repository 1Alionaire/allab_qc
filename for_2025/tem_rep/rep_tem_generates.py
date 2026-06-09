import json
from pathlib import Path
import math
import random
import copy
import os
import sys
import logging


BASE_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE_DIR))
from utils.utilities import lab_id_sort_key, tem_calc_asb_percent, replicate_analyst, calc_asb_percent

logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s'
)

def get_plm_nob_random_sample_from_project(samples_list, analyst): 
    samples_list_duplicate = copy.deepcopy(samples_list)

    while True:
        random_index = random.randint(0, len(samples_list_duplicate) - 1)
        removed_element = samples_list_duplicate.pop(random_index)
        logging.info('=' * 50)
        logging.info(removed_element)
        if (removed_element['Asb Type Type 1'] != 'PS' 
            and len(removed_element['Lab ID']) > 3
            and removed_element['Client ID'] != 'R' 
            and removed_element['Lab ID'] != 'None'
            and removed_element['Client ID'][-2:] not in ['AB', 'KK', 'OV', 'VC']):
            original_sample = removed_element
            break

    duplicate_sample = copy.deepcopy(original_sample)        #random_calc_point_asb(samples_list[i][f'Point {j}'])

    duplicate_sample['Client ID'] = 'R' + original_sample['Client ID'] + replicate_analyst(analyst)

    tem_calc_asb_dict = tem_calc_asb_percent(original_sample)
    
    if duplicate_sample['Asb Type Type 1'] != 'NAD' and duplicate_sample['Asb Type Type 1'] != 'None':
        duplicate_sample['Percent Type 1'] = tem_calc_asb_dict['percent']
        duplicate_sample['Point Type 1'] = tem_calc_asb_dict['point']

    if duplicate_sample['Asb Type Type 1'] == 'None':
        duplicate_sample['Asb Type Type 1'] = 'NAD'
        duplicate_sample['Percent Type 1'] = 'NAD'
        
        original_sample['Asb Type Type 1'] = 'NAD'
        original_sample['Percent Type 1'] = 'NAD'


    return {
        'sample': original_sample['Client ID'],
        '1st_analyst_asb_type': original_sample['Asb Type Type 1'], 
        '1st_analyst_asb_percent': original_sample['Percent Type 1'], 
        'dup_asb_type': duplicate_sample['Asb Type Type 1'], 
        'dup_asb_percent': duplicate_sample['Percent Type 1'],
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
                tem_total_samples = 0       # total count samples per day. 
                # filter all projects by day
                projects_by_date = {key: value for key, value in raw_data.items() if (value['date'][:-8].strip() == record['date'] 
                                                                                      and value['tem_count'] > 0) 
                                                                                      and value['analyst'] == analyst }
                tem_counter  = 0         # this counter need only for first blank sample in first project of the day
                #iterate through all projects in this day
                for project, project_info in projects_by_date.items():
                    
                    if project_info['tem_count'] > 0 and project_info['tem_analysis']:
                        before = tem_total_samples
                        after = before + project_info['tem_count']

                        blanks_before = before // 100
                        blanks_after = after // 100
                        blanks_to_add = blanks_after - blanks_before

                        duplicates_before = before // 15
                        duplicates_after = after // 15
                        duplicates_to_add = duplicates_after - duplicates_before

                        if tem_counter == 0:
                            duplicates_to_add += 1
                            blanks_to_add += 1

                        if duplicates_to_add > 0:
                            for i in range(duplicates_to_add):
                                sample_duplicate = get_plm_nob_random_sample_from_project(project_info['tem_analysis'], analyst)

                                duplicate_record = {
                                    'date': project_info['date'], 
                                    'project': project,
                                    'lab id': sample_duplicate['whole_original']['Lab ID'],
                                    'file_name': project_info['file_name'],
                                    'sample': sample_duplicate['sample'],
                                    '1st_analyst_name': project_info['analyst'],
                                    '1st_analyst_asb_type': sample_duplicate['1st_analyst_asb_type'],
                                    '1st_analyst_asb_percent': sample_duplicate['1st_analyst_asb_percent'],
                                    'dup_name': sample_duplicate['whole_duplicate']['Client ID'][-2:], 
                                    'dup_asb_type': sample_duplicate['dup_asb_type'],
                                    'dup_asb_percent': sample_duplicate['dup_asb_percent'],
                                    'whole_original': sample_duplicate['whole_original'],
                                    'whole_duplicate': sample_duplicate['whole_duplicate']
                                }
                                
                                data_per_analyst.append(duplicate_record)
                                data_per_month['total'].append(duplicate_record)

                                for l in project_info['tem_analysis']:
                                    if l['Client ID'] == duplicate_record['sample']:
                                        project_info['tem_analysis'].remove(l)

                        if blanks_to_add > 0:
                            for i in range(blanks_to_add):
                                duplicate_record = {
                                    'date': project_info['date'], 
                                    'project': project,
                                    'lab id': project + '-1000',
                                    'file_name': project_info['file_name'],
                                    'sample': 'bl',
                                    '1st_analyst_name': project_info['analyst'],
                                    'dup_name': replicate_analyst(project_info['analyst']) , 
                                    }
                                data_per_analyst.append(duplicate_record)
                                data_per_month['total'].append(duplicate_record)
                    
                        tem_total_samples = after
                        tem_counter+=1

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