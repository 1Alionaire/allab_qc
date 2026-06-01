import json
from pathlib import Path
import math
import random
import copy

# prod data
# analysts = ['OV', 'KK', 'AB', 'VC']

analysts = ['VC']


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

    # print(f'samples_list: {samples_list}')
    # print(f'type: {type}')
    # print(f'samples_list: {samples_list}')

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
            # TO DO: replace key Percent 1 Option for Type Asb 1 Option
            duplicate_sample['Percent 1 Option'] = round(((intermediate_percent_result * float(duplicate_sample['Total Residue'])) / 100), 2) #change

    return duplicate_sample

def main_start():
    result_to_export = {}

    # for mac
    grouped_date_file_path = Path(str(Path.cwd()) + '/analysts_grouped.json')
    output_file_path = Path(str(Path.cwd()) + '/all_analyst_dup_data.json')
    # for windows:
    # grouped_date_file_path = Path(r"C:\Python\allab\allab_qc\data\analysts_grouped.json")
    # output_file_path = Path(fr"C:\Python\allab\allab_qc\data\all_analyst_dup_data.json")

    with grouped_date_file_path.open("r", encoding="utf-8") as f:
        grouped_data_analysts = json.load(f) 
    
    for enter_analyst in analysts:
        result_to_export[enter_analyst] = make_samples_by_analyst(enter_analyst, grouped_data_analysts)

    with open(output_file_path, "a", encoding="utf-8") as file:
        json.dump(result_to_export, file, indent=2, ensure_ascii=False)


def make_samples_by_analyst(analyst, grouped_data_analysts):
    final_result = {}

    # for mac
    path = Path(str(Path.cwd()) + f'/{analyst}_data.jsonl')
    # for windows
    # path = Path(fr"C:\Python\allab\allab_qc\data\{analyst}_data.jsonl")
    

    with path.open("r", encoding="utf-8") as f:
        data = json.load(f)       

    grouped_data_analysts = grouped_data_analysts[analyst]

    for record in grouped_data_analysts:
        # how many blank samples per day accroding audit
        limit = 100
        # total count samples per day. 
        nob_total_samples = 0
        plm_total_samples = 0
        tem_total_samples = 0
        # filter all projects by day
        projects_by_date = {key: value for key, value in data.items() if value['date'] == record['date'] }
        # this counter need only for first blank sample in first project of the day
        counter = 0
        plm_counter = 0


        #iterate through all projects in this day
        for project, project_info in projects_by_date.items():
            final_project_info = {'plm_rep_samples': [], 'plm_dup_samples': [], 
                            'nob_rep_samples': [], 'nob_dup_samples': [],
                            'tem_rep_samples': [], 'tem_dup_samples': [], 
                            'file_name':project_info['file_name'], 'analyst':project_info['analyst']}

            if project_info['plm_count'] > 0:
                before = plm_total_samples
                after = before + project_info['plm_count']

                # counting replicates
                replicates_before = before // 15
                replicates_after = after // 15
                replicates_to_add = replicates_after - replicates_before

                duplicates_before = before // 50
                duplicates_after = after // 50
                duplicates_to_add = duplicates_after - duplicates_before

                if plm_counter == 0:
                    replicates_to_add += 1
                    duplicates_to_add += 1

                if duplicates_to_add > 0:
                    for i in range(duplicates_to_add):
                        final_project_info['plm_dup_samples'].append(get_plm_nob_random_sample_from_project(project_info['plm_analysis'], 'dup', project_info['analyst']))

                if replicates_to_add > 0:
                    for i in range(replicates_to_add):
                        final_project_info['plm_rep_samples'].append(get_plm_nob_random_sample_from_project(project_info['plm_analysis'], 'rep', project_info['analyst']))

                

                final_project_info['rep_plm_count'] = replicates_to_add
                final_project_info['dup_plm_count'] = duplicates_to_add
                plm_total_samples = after
                plm_counter+=1

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

                if duplicates_to_add > 0:
                    for i in range(duplicates_to_add):
                        final_project_info['nob_dup_samples'].append(get_plm_nob_random_sample_from_project(project_info['nob_analysis'], 'dup', project_info['analyst']))

                if replicates_to_add > 0:
                    for i in range(replicates_to_add):
                        final_project_info['nob_rep_samples'].append(get_plm_nob_random_sample_from_project(project_info['nob_analysis'], 'rep', project_info['analyst']))

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
                                "Type Asb 1 Option": "NAD",
                                "Percent 1 Option": "NAD",
                                "Type Asb 2 Option": "",
                                "Percent 2 Option": "",
                                "Vermiculite": "ND",
                                "Method": "198.6",
                                "Undesolved Materials": "",
                                "Total Residue": ""
                            })
  

                final_project_info['blank_nob_count'] = blanks_to_add
                final_project_info['rep_nob_count'] = replicates_to_add
                final_project_info['dup_nob_count'] = duplicates_to_add
                nob_total_samples = after
                counter+=1
            
            if project_info['tem_count'] > 0:
                # general
                before = tem_total_samples
                after = before + project_info['tem_count']

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

                if duplicates_to_add > 0:
                    for i in range(duplicates_to_add):
                        final_project_info['tem_dup_samples'].append(get_plm_nob_random_sample_from_project(project_info['nob_analysis'], 'dup', project_info['analyst']))

                if replicates_to_add > 0:
                    for i in range(replicates_to_add):
                        final_project_info['tem_rep_samples'].append(get_plm_nob_random_sample_from_project(project_info['nob_analysis'], 'rep', project_info['analyst']))

                if blanks_to_add > 0:
                    for i in range(blanks_to_add):
                        final_project_info['nob_rep_samples'].append({
                               "Client ID": "BL",
                                "Lab ID": "260115-22-6",
                                "Layer": 1,
                                "Homogeneity": "C",
                                "Residue": 1.91,
                                "Point Type 1": "NAD",
                                "Percent Type 1": "NAD",
                                "Asb Type Type 1": '',
                                "Point Type 2": "NAD",
                                "Percent Type 2": "NAD",
                                "Asb Type Type 2": '',
                                "Microscope ": "Y",
                                "Eccentricity ": "Y",
                                "Grid Pre": "Y",
                                "Grid Box #": '',
                                "Grid Box ID 1": "D2",
                                "Grid Box ID 2": "E2",
                                "Method": 198.6,
                                "NA or PS": ''
            
                            })
  

                final_project_info['blank_nob_count'] = blanks_to_add
                final_project_info['rep_nob_count'] = replicates_to_add
                final_project_info['dup_nob_count'] = duplicates_to_add

                tem_total_samples = after
                counter+=1
 
            
            final_result[project] = final_project_info

    return final_result

if __name__ == "__main__":
    main_start()