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

massive_lab_id = []

script_dir = Path(__file__).resolve().parent

for file in script_dir.glob("*.json"):
    with file.open("r", encoding="utf-8") as f:
        data_from_file = json.load(f)

    data_from_file = {key : value for key, value in data_from_file.items() if key != 'total' }
    for analyst, info in data_from_file.items():
        for item in info:
            # logging.info(item['lab id'])
            massive_lab_id.append(item['lab id'])

massive_lab_id_unique = set(massive_lab_id)
print(len(massive_lab_id_unique))
print(len(massive_lab_id))
if len(massive_lab_id_unique) == len(massive_lab_id):
    print('No duplicates')