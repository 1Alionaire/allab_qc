import threading
from pathlib import Path
from tkinter import Tk, Button, Label, Text, filedialog, messagebox, END, DISABLED, NORMAL
from tkinter.ttk import Progressbar
from openpyxl import Workbook, load_workbook
import pandas as pd
from datetime import date
import json
from tkcalendar import DateEntry
from datetime import datetime
import json
import math
import random
import copy
import os
import logging

logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s'
)

def replicate_analyst(entry_analyst):
    if entry_analyst == "AB":
        return 'KK'
    elif entry_analyst == "KK":
        return 'AB'
    elif entry_analyst == "OV":
        return 'AB'
    elif entry_analyst == "VC":
        return 'AB'
    elif entry_analyst == "No Analyst":
        return 'No Analyst'

def lab_id_sort_key(item):
    lab_id = str(item.get("lab id", ""))
    lab_id = lab_id.replace(' ', '')
    try:
        main_part, number_part, sample_part = lab_id.split("-")
        return int(main_part), int(number_part), int(sample_part)
    except ValueError:
        return 999999999, 999999999
    
def random_calc_point_asb(inp_str):
    if inp_str == 'None':
        return 'None'
    if inp_str == '50':
        return '50'
    operation = random.choice(['-', '+'])
    if operation == '-':
        value = random.choice([1, 2, 3, 4])
        if (int(inp_str) - value) > 2:
            return str(int(inp_str) - value)
        else:
            return '2'
    else:
        value = random.choice([1, 2, 3, 4])
        if (int(inp_str) + value) < 50:
            return str(int(inp_str) + value)
        else:
            return '49'
        
def calc_std_dev(first_inp, second_inp):
    if first_inp == 'NAD' and second_inp == 'NAD':
        return 0
    try:
        float_first_inp = float(first_inp)
    except:
        float_first_inp = 0
    try:
        float_second_inp = float(second_inp)
    except:
        float_second_inp = 0
        
    if float_first_inp != 0 and float_second_inp != 0:
        return round(abs((float_first_inp - float_second_inp) / ((float_first_inp + float_second_inp) / 2)) , 2)
    else:
        return 0

def calc_asb_percent(duplicate_sample):
    sum_points = 0
    add_asb_glass = 0

    for j in range(1, 9):
        if duplicate_sample[f'Point {j}'] == 'None':
            pass
        elif duplicate_sample[f'Point {j}'] == '50':
            sum_points += 50
        else:
            try: 
                sum_points += int(duplicate_sample[f'Point {j}'])
                add_asb_glass += 1
            except:
                pass

    return round((((add_asb_glass) / sum_points) * 100), 2)  
       
def tem_calc_asb_percent(duplicate_sample):
    original_point = duplicate_sample["Point Type 1"]

    if original_point == 'None':
        return 'None'
    
    int_original_point = 0
    try:
        int_original_point = int(original_point)
    except:
        int_original_point = 0

    operation = random.choice(['-', '+'])
    value = random.choice([2, 3, 4])

    if int_original_point != 0:
        if operation == '-':
            inter_result = int_original_point - value
            if inter_result < 0: 
                inter_result = 1
        else:
            inter_result = int_original_point + value
    else:
        return  {'percent': 'NAD', 
                'point': 'NAD'}

    try:
        float_residue = float(duplicate_sample["Residue"])
    except:
        return  {'percent': 0, 
                'point': inter_result}

    return {'percent': round(((float_residue * inter_result) / 100), 2), 
            'point': inter_result}   


def get_plm_rep_sample_from_project(samples_list, analyst):
    rep_analyst = replicate_analyst(analyst) 
    samples_list_duplicate = copy.deepcopy(samples_list)

    while True:
        random_index = random.randint(0, len(samples_list_duplicate) - 1)
        removed_element = samples_list_duplicate.pop(random_index)
        if removed_element['Type Asb 1 Option'] != 'PS':
            original_sample = removed_element
            break

    duplicate_sample = copy.deepcopy(original_sample) 

    duplicate_sample['Client ID'] = 'R' + original_sample['Client ID'] + rep_analyst

    if duplicate_sample['Type Asb 1 Option'] != 'NAD':
        for j in range(1, 9):
            if duplicate_sample[f'Type {j}'] != 'None':
                duplicate_sample[f'Point {j}'] = random_calc_point_asb(original_sample[f'Point {j}'])

        intermediate_percent_result = calc_asb_percent(duplicate_sample)

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

def get_plm_dup_sample_from_project(samples_list, analyst):
    samples_list_duplicate = copy.deepcopy(samples_list)

    while True:
        random_index = random.randint(0, len(samples_list_duplicate) - 1)
        removed_element = samples_list_duplicate.pop(random_index)
        if removed_element['Type Asb 1 Option'] != 'PS':
            original_sample = removed_element
            break

    duplicate_sample = copy.deepcopy(original_sample) 

    duplicate_sample['Client ID'] = 'D' + original_sample['Client ID'] + analyst

    if duplicate_sample['Type Asb 1 Option'] != 'NAD':
        for j in range(1, 9):
            if duplicate_sample[f'Type {j}'] != 'None':
                duplicate_sample[f'Point {j}'] = random_calc_point_asb(original_sample[f'Point {j}'])

        intermediate_percent_result = calc_asb_percent(duplicate_sample)

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

def get_nob_rep_random_sample_from_project(samples_list, analyst): 
    rep_analyst = replicate_analyst(analyst) 
    samples_list_duplicate = copy.deepcopy(samples_list)

    while True:
        random_index = random.randint(0, len(samples_list_duplicate) - 1)
        removed_element = samples_list_duplicate.pop(random_index)
        if removed_element['Type Asb 1 Option'] != 'PS':
            if removed_element['Method'] != 'None':
                original_sample = removed_element
                break

    duplicate_sample = copy.deepcopy(original_sample)        #random_calc_point_asb(samples_list[i][f'Point {j}'])

    duplicate_sample['Client ID'] = 'R' + original_sample['Client ID'] + rep_analyst

    if duplicate_sample['Type Asb 1 Option'] != 'NAD' and duplicate_sample['Type Asb 1 Option'] != 'None':
        for j in range(1, 9):
            if duplicate_sample[f'Type {j}'] != 'None':
                duplicate_sample[f'Point {j}'] = random_calc_point_asb(original_sample[f'Point {j}'])

        intermediate_percent_result = calc_asb_percent(duplicate_sample)

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

def get_nob_dup_random_sample_from_project(samples_list, analyst): 
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

        intermediate_percent_result = calc_asb_percent(duplicate_sample)

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

def get_tem_rep_random_sample_from_project(samples_list, analyst): 
    samples_list_duplicate = copy.deepcopy(samples_list)
    while True:
        random_index = random.randint(0, len(samples_list_duplicate) - 1)
        removed_element = samples_list_duplicate.pop(random_index)
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

def get_tem_dup_random_sample_from_project(samples_list, analyst): 
    samples_list_duplicate = copy.deepcopy(samples_list)
    while True:
        random_index = random.randint(0, len(samples_list_duplicate) - 1)
        removed_element = samples_list_duplicate.pop(random_index)
        if (removed_element['Asb Type Type 1'] != 'PS' 
            and len(removed_element['Lab ID']) > 3
            and removed_element['Client ID'] != 'R' 
            and removed_element['Lab ID'] != 'None'
            and removed_element['Client ID'][-2:] not in ['AB', 'KK', 'OV', 'VC']):
            original_sample = removed_element
            break

    duplicate_sample = copy.deepcopy(original_sample)  

    duplicate_sample['Client ID'] = 'D' + original_sample['Client ID'] + analyst

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



def get_plm_replicate_samples(projects):
    plm_projects = {key: value for key, value in projects.items() if value['plm_count'] > 0}
    samples = []

    for key, value in plm_projects.items():
        plm_replicates_to_add = math.ceil(value['plm_count'] / 15)

        for i in range(plm_replicates_to_add):
            sample_duplicate = get_plm_rep_sample_from_project(value['plm_analysis'], value['analyst'])
            duplicate_record = {
                            'type' : 'plm_rep',
                            'date': value['date'], 
                            'project': key,
                            'lab id': sample_duplicate['whole_original']['Lab ID'],
                            'file_name': value['file_name'],
                            'sample': sample_duplicate['sample'],
                            '1st_analyst_name': value['analyst'],
                            '1st_analyst_asb_type': sample_duplicate['1st_analyst_asb_type'],
                            '1st_analyst_asb_percent': sample_duplicate['1st_analyst_asb_percent'],
                            'dup_name': sample_duplicate['whole_duplicate']['Client ID'][-2:], 
                            'dup_asb_type': sample_duplicate['dup_asb_type'],
                            'dup_asb_percent': sample_duplicate['dup_asb_percent'],
                            'whole_original': sample_duplicate['whole_original'],
                            'whole_duplicate': sample_duplicate['whole_duplicate']
                        }
            samples.append(duplicate_record)   

    return samples

def get_plm_duplicate_samples(projects):
    plm_total_samples = 0
    plm_projects = {key: value for key, value in projects.items() if value['plm_count'] > 0}
    samples = []

    for key, value in plm_projects.items():

        before = plm_total_samples
        after = before + value['plm_count']
        plm_duplicates_before = math.ceil(before / 40)  # before // 50
        plm_duplicates_after = math.ceil(after / 40)  # after // 50
        plm_duplicates_to_add = plm_duplicates_after - plm_duplicates_before

        if plm_duplicates_to_add > 0:
            for i in range(plm_duplicates_to_add):
                sample_duplicate = get_plm_dup_sample_from_project(value['plm_analysis'], value['analyst'])
                duplicate_record = {
                                'type' : 'plm_dup',
                                'date': value['date'], 
                                'project': key,
                                'lab id': sample_duplicate['whole_original']['Lab ID'],
                                'file_name': value['file_name'],
                                'sample': sample_duplicate['sample'],
                                '1st_analyst_name': value['analyst'],
                                '1st_analyst_asb_type': sample_duplicate['1st_analyst_asb_type'],
                                '1st_analyst_asb_percent': sample_duplicate['1st_analyst_asb_percent'],
                                'dup_name': sample_duplicate['whole_duplicate']['Client ID'][-2:], 
                                'dup_asb_type': sample_duplicate['dup_asb_type'],
                                'dup_asb_percent': sample_duplicate['dup_asb_percent'],
                                'whole_original': sample_duplicate['whole_original'],
                                'whole_duplicate': sample_duplicate['whole_duplicate']
                            }
                samples.append(duplicate_record)

        plm_total_samples = after        

    return samples

def get_nob_replicate_samples(projects):
    nob_total_samples = 0
    nob_projects = {key: value for key, value in projects.items() if value['nob_count'] > 0}
    samples = []

    for key, value in nob_projects.items():
        before = nob_total_samples
        after = before + value['nob_count']

        blanks_before = math.ceil(before / 80)  # before // 100
        blanks_after = math.ceil(after / 80)  # after // 100
        blanks_to_add = blanks_after - blanks_before

        nob_replicates_to_add = math.ceil(value['nob_count'] / 15)

        for i in range(nob_replicates_to_add):
            sample_duplicate = get_nob_rep_random_sample_from_project(value['nob_analysis'], value['analyst'])
            duplicate_record = {
                            'type' : 'nob_rep',
                            'date': value['date'], 
                            'project': key,
                            'lab id': sample_duplicate['whole_original']['Lab ID'],
                            'file_name': value['file_name'],
                            'sample': sample_duplicate['sample'],
                            '1st_analyst_name': value['analyst'],
                            '1st_analyst_asb_type': sample_duplicate['1st_analyst_asb_type'],
                            '1st_analyst_asb_percent': sample_duplicate['1st_analyst_asb_percent'],
                            'dup_name': sample_duplicate['whole_duplicate']['Client ID'][-2:], 
                            'dup_asb_type': sample_duplicate['dup_asb_type'],
                            'dup_asb_percent': sample_duplicate['dup_asb_percent'],
                            'whole_original': sample_duplicate['whole_original'],
                            'whole_duplicate': sample_duplicate['whole_duplicate']
                        }
            samples.append(duplicate_record)

        if blanks_to_add > 0:
            for i in range(blanks_to_add):
                duplicate_record = {
                    'type' : 'nob_rep',
                    'date': value['date'], 
                    'project': key,
                    'lab id': key + '-1000',
                    'file_name': value['file_name'],
                    'sample': 'bl',
                    '1st_analyst_name': value['analyst'],
                    'dup_name': replicate_analyst(value['analyst']) , 
                    }
                samples.append(duplicate_record)

        nob_total_samples = after        
    return samples

def get_nob_duplicate_samples(projects):
    nob_total_samples = 0
    nob_projects = {key: value for key, value in projects.items() if value['nob_count'] > 0}
    samples = []

    for key, value in nob_projects.items():
        before = nob_total_samples
        after = before + value['nob_count']

        duplicates_before = math.ceil(before / 40)  # before // 50
        duplicates_after = math.ceil(after / 40)  # after // 50
        duplicates_to_add = duplicates_after - duplicates_before

        if duplicates_to_add > 0:
            for i in range(duplicates_to_add):
                sample_duplicate = get_nob_dup_random_sample_from_project(value['nob_analysis'], value['analyst'])
                duplicate_record = {
                                'type' : 'nob_dup',
                                'date' : value['date'], 
                                'project' : key,
                                'lab id': sample_duplicate['whole_original']['Lab ID'],
                                'file_name': value['file_name'],
                                'sample': sample_duplicate['sample'],
                                '1st_analyst_name': value['analyst'],
                                '1st_analyst_asb_type': sample_duplicate['1st_analyst_asb_type'],
                                '1st_analyst_asb_percent': sample_duplicate['1st_analyst_asb_percent'],
                                'dup_name': sample_duplicate['whole_duplicate']['Client ID'][-2:], 
                                'dup_asb_type': sample_duplicate['dup_asb_type'],
                                'dup_asb_percent': sample_duplicate['dup_asb_percent'],
                                'whole_original': sample_duplicate['whole_original'],
                                'whole_duplicate': sample_duplicate['whole_duplicate']
                            }
                samples.append(duplicate_record)
        nob_total_samples = after        

    return samples

def get_tem_replicate_samples(projects):
    tem_total_samples = 0
    tem_projects = {key: value for key, value in projects.items() if value['tem_count'] > 0}
    samples = []

    for key, value in tem_projects.items():
        before = tem_total_samples
        after = before + value['tem_count']

        blanks_before = math.ceil(before / 80)  # before // 100
        blanks_after = math.ceil(after / 80)  # after // 100
        blanks_to_add = blanks_after - blanks_before
        
        replicates_to_add =  math.ceil(value['tem_count'] / 15)

        for i in range(replicates_to_add):
            sample_duplicate = get_tem_rep_random_sample_from_project(value['tem_analysis'], value['analyst'])
            duplicate_record = {
                            'type' : 'tem_rep',
                            'date': value['date'], 
                            'project': key,
                            'lab id': sample_duplicate['whole_original']['Lab ID'],
                            'file_name': value['file_name'],
                            'sample': sample_duplicate['sample'],
                            '1st_analyst_name': value['analyst'],
                            '1st_analyst_asb_type': sample_duplicate['1st_analyst_asb_type'],
                            '1st_analyst_asb_percent': sample_duplicate['1st_analyst_asb_percent'],
                            'dup_name': sample_duplicate['whole_duplicate']['Client ID'][-2:], 
                            'dup_asb_type': sample_duplicate['dup_asb_type'],
                            'dup_asb_percent': sample_duplicate['dup_asb_percent'],
                            'whole_original': sample_duplicate['whole_original'],
                            'whole_duplicate': sample_duplicate['whole_duplicate']
                        }
            samples.append(duplicate_record)

        if blanks_to_add > 0:
            for i in range(blanks_to_add):
                duplicate_record = {
                    'type' : 'tem_rep',
                    'date': value['date'], 
                    'project': key,
                    'lab id': key + '-1000',
                    'file_name': value['file_name'],
                    'sample': 'bl',
                    '1st_analyst_name': value['analyst'],
                    'dup_name': replicate_analyst(value['analyst']) , 
                    }
                samples.append(duplicate_record)
        tem_total_samples = after        

    return samples

def get_tem_duplicate_samples(projects):
    tem_total_samples = 0
    tem_projects = {key: value for key, value in projects.items() if value['tem_count'] > 0}
    samples = []

    for key, value in tem_projects.items():
        before = tem_total_samples
        after = before + value['tem_count']

        duplicates_before = math.ceil(before / 40)  # before // 50
        duplicates_after = math.ceil(after / 40) # after // 50 
        duplicates_to_add = duplicates_after - duplicates_before

        if duplicates_to_add > 0:
            for i in range(duplicates_to_add):
                sample_duplicate = get_tem_dup_random_sample_from_project(value['tem_analysis'], value['analyst'])
                duplicate_record = {
                                'type' : 'tem_dup',
                                'date' : value['date'], 
                                'project' : key,
                                'lab id': sample_duplicate['whole_original']['Lab ID'],
                                'file_name': value['file_name'],
                                'sample': sample_duplicate['sample'],
                                '1st_analyst_name': value['analyst'],
                                '1st_analyst_asb_type': sample_duplicate['1st_analyst_asb_type'],
                                '1st_analyst_asb_percent': sample_duplicate['1st_analyst_asb_percent'],
                                'dup_name': sample_duplicate['whole_duplicate']['Client ID'][-2:], 
                                'dup_asb_type': sample_duplicate['dup_asb_type'],
                                'dup_asb_percent': sample_duplicate['dup_asb_percent'],
                                'whole_original': sample_duplicate['whole_original'],
                                'whole_duplicate': sample_duplicate['whole_duplicate']
                            }
                samples.append(duplicate_record)
        tem_total_samples = after        

    return samples


def generate_duplicates(input_data):
    analysts_list = []

    for key, value in input_data.items():              
        analysts_list.append(value['analyst'])
    
    unique_analysts_list = list(set(analysts_list))

    total_data = []
    for analyst in unique_analysts_list:
        
        analysts_projects = {key: value for key, value in input_data.items() if value['analyst'] == analyst}

        plm_replicate_samples_list = get_plm_replicate_samples(analysts_projects)
        if plm_replicate_samples_list:
            for i in plm_replicate_samples_list:
                total_data.append(i)
        
        plm_duplicate_samples_list = get_plm_duplicate_samples(analysts_projects)
        if plm_duplicate_samples_list:
            for i in plm_duplicate_samples_list:
                total_data.append(i)

        nob_replicate_samples_list = get_nob_replicate_samples(analysts_projects)
        if nob_replicate_samples_list:
            for i in nob_replicate_samples_list:
                total_data.append(i)
        
        nob_duplicate_samples_list = get_nob_duplicate_samples(analysts_projects)
        if nob_duplicate_samples_list:
            for i in nob_duplicate_samples_list:
                total_data.append(i)

        tem_replicate_samples_list = get_tem_replicate_samples(analysts_projects)
        if tem_replicate_samples_list:
            for i in tem_replicate_samples_list:
                total_data.append(i)

        tem_duplicate_samples_list = get_tem_duplicate_samples(analysts_projects)
        if tem_duplicate_samples_list:
            for i in tem_duplicate_samples_list:
                total_data.append(i)

    
    return total_data
        


