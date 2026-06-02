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
        return str(int(inp_str) - value)
    else:
        value = random.choice([1, 2, 3])
        if (int(inp_str) + value) < 50:
            return str(int(inp_str) + value)
        else:
            return '49'