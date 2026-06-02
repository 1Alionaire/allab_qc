import json
from pathlib import Path
import math
import random
import copy
import os

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