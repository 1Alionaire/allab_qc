import json
from pathlib import Path
import math
import random


def get_random_sample_from_project(projects_list):
    return random.choice(projects_list)


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
    plm_total_samples = 0
    # filter all projects by day
    projects_by_date = {key: value for key, value in data.items() if value['date'] == record['date'] }
    # this counter need only for first blank sample in first project of the day
    counter = 0


    #iterate through all projects in this day
    for project, project_info in projects_by_date.items():
        final_project_info = {'plm_rep_samples': [], 'plm_dup_samples': [], 
                        'nob_rep_samples': [], 'nob_dup_samples': [],
                        'tem_rep_samples': [], 'tem_dup_samples': [],}

        if project_info['plm_count'] > 0:
            duplicate_samples_amount = math.ceil(project_info['plm_count'] / 50)
            for i in range(duplicate_samples_amount):
                final_project_info['plm_dup_samples'].append(get_random_sample_from_project(project_info['plm_analysis']))

            replicate_samples_amount = math.ceil(project_info['plm_count'] / 15)
            for i in range(replicate_samples_amount):
                final_project_info['plm_rep_samples'].append(get_random_sample_from_project(project_info['plm_analysis']))


        if project_info['nob_count'] > 0:
            duplicate_samples_amount = math.ceil(project_info['nob_count'] / 50)
            for i in range(duplicate_samples_amount):
                final_project_info['nob_dup_samples'].append(get_random_sample_from_project(project_info['nob_analysis']))

            replicate_samples_amount = math.ceil(project_info['nob_count'] / 15)
            for i in range(replicate_samples_amount):
                final_project_info['nob_rep_samples'].append(get_random_sample_from_project(project_info['nob_analysis']))

            
            # sum samples before adding samples from current project
            before = plm_total_samples
            # sum samples after adding samples from current project
            after = before + project_info['nob_count']
            # divide by samples amount how many samples before and after 
            blanks_before = before // limit
            blanks_after = after // limit
            # how many blanks? difference between before and after
            blanks_to_add = blanks_after - blanks_before
            # if first project of the day - it must be one blank
            if counter == 0:
                blanks_to_add = 1
            #add result
            final_project_info['blank_nob_count'] = blanks_to_add
            # new total amount samples per day
            plm_total_samples = after
            # increase counter for day
            counter+=1
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


        final_result[project] = final_project_info


# print(final_result)

with open(output_file, "w", encoding="utf-8") as file:
    json.dump(final_result, file, indent=4, ensure_ascii=False)