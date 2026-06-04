import json
from pathlib import Path
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
        value = random.choice([1, 2, 3])
        if (int(inp_str) - value) > 0:
            return str(int(inp_str) - value)
        else:
            return '1'
    else:
        value = random.choice([1, 2, 3])
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

def return_excel_filename(type, input_filename):
    if type == 'PLM':
        if '_2026-01' in str(input_filename):
            return 'PLM_Jan.xlsx'
        elif '_2026-02' in str(input_filename):
            return 'PLM_Feb.xlsx'
        elif '_2026-03' in str(input_filename):
            return 'PLM_Mar.xlsx'
        elif '_2026-04' in str(input_filename):
            return 'PLM_Apr.xlsx'
        elif '_2026-05' in str(input_filename):
            return 'PLM_May.xlsx'
    elif type == 'NOB':
        if '_2026-01' in str(input_filename):
            return 'NOB_Jan.xlsx'
        elif '_2026-02' in str(input_filename):
            return 'NOB_Feb.xlsx'
        elif '_2026-03' in str(input_filename):
            return 'NOB_Mar.xlsx'
        elif '_2026-04' in str(input_filename):
            return 'NOB_Apr.xlsx'
        elif '_2026-05' in str(input_filename):
            return 'NOB_May.xlsx'
    elif type == 'TEM':
        if '_2026-01' in str(input_filename):
            return 'TEM_Jan.xlsx'
        elif '_2026-02' in str(input_filename):
            return 'TEM_Feb.xlsx'
        elif '_2026-03' in str(input_filename):
            return 'TEM_Mar.xlsx'
        elif '_2026-04' in str(input_filename):
            return 'TEM_Apr.xlsx'
        elif '_2026-05' in str(input_filename):
            return 'TEM_May.xlsx'
        
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
    
    logging.info('*' * 40)
    logging.info(f"original_sample['Lab ID'] : {duplicate_sample['Lab ID']}")
    logging.info(f'inter_result: {inter_result}')
    logging.info(f'float_residue: {float_residue}')
    logging.info(f'(float_residue * inter_result): {(float_residue * inter_result)}')
    return {'percent': round(((float_residue * inter_result) / 100), 2), 
            'point': inter_result}
    