import sys
import json
import logging
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from pathlib import Path

import win32com.client as win32
import pythoncom

logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s'
)


def get_resource_path(name):
    """Путь к ресурсу — внутри .exe или рядом с .py."""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / name
    return Path(__file__).resolve().parent / name


def get_external_path(name):
    """Путь к внешнему файлу (рядом с .exe или .py)."""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent / name
    return Path(__file__).resolve().parent / name

OPTIONS = ["PLM_REP", "PLM_DUP", "NOB_REP", "NOB_DUP"]

plm_columns_dict = {
    1: "Client ID",
    2: "Lab ID",
    3: "Layer",
    4: "Color",
    5: "Texture",
    6: "Homogeneity",
    7: "Morphology",
    8: "RI II Type 1",
    9: "RI II Type 2",
    10: "RI ┴ Type 1",
    11: "RI ┴ Type 2",
    12: "Sign of \nElongation Type 1",
    13: "Sign of \nElongation Type 2",
    14: "Extinction \nAngle Type 1",
    15: "Extinction \nAngle Type 2",
    16: "Pleochroism /\nColor Type 1",
    17: "Pleochroism /\nColor Type 2",
    18: "Birefringence Type 1",
    19: "Birefringence Type 2",
    20: "Other Fibers",
    21: "Property",
    22: "% Non-\nAsbestos",
    23: "Type 1",
    24: "Point 1",
    25: "Type 2",
    26: "Point 2",
    27: "Type 3",
    28: "Point 3",
    29: "Type 4",
    30: "Point 4",
    31: "Type 5",
    32: "Point 5",
    33: "Type 6",
    34: "Point 6",
    35: "Type 7",
    36: "Point 7",
    37: "Type 8",
    38: "Point 8",
    39: "Type Asb 1 Option",
    40: "Percent 1 Option",
    41: "Type Asb 2 Option",
    42: "Percent 2 Option",
    43: "Vermiculite",
    44: "Method",
    45: "Undesolved Materials",
    46: "Total Residue",
}

tem_columns_dict = {
    1: "Client ID",
    2: "Lab ID",
    3: "Layer",
    4: "Homogeneity",
    5: 'Residue', 
    6: 'Point Type 1', 
    7: 'Percent Type 1', 
    8: 'Asb Type Type 1', 
    9: 'Point Type 2', 
    10: 'Percent Type 2',
    11: 'Asb Type Type 2', 
    12: 'Microscope', 
    13: 'Eccentricity', 
    14: 'Grid Pre', 
    15: 'Grid Box #',
    16: 'Grid Box ID 1', 
    17: 'Grid Box ID 2', 
    18: 'Method', 
    19: 'NA or PS'
}

def fill_all_excel_files(inp_data):
    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1  # без Protected View

    processed = 0
    errors = 0
    try:
        for sample in inp_data:
            fname = sample.get("file_name", "")
            logging.info(fname)

            wb = None
            try:
                wb = excel.Workbooks.Open(str(Path(fname).resolve()))

                type_analysis = sample.get("type", "")
                print(f'project: {sample.get("project", "")}')
                print(f'type_analysis: {type_analysis}')
                if ('nob' in type_analysis) or ('plm' in type_analysis):
                    plm_analysis_ws = wb.Worksheets("SampleAnalyses")
                    count = 8
                    last_sample_count = 0
                    while True:
                        value = plm_analysis_ws.Range(f"B{count}").Value
                        if value is not None and str(value).strip() != "":
                            count += 1
                        else:
                            last_sample_count = count
                            break

                    if sample.get("whole_duplicate"):
                        for col, text_key in plm_columns_dict.items():
                            v = sample["whole_duplicate"].get(text_key, "None")
                            plm_analysis_ws.Cells(last_sample_count, col).Value = "" if v == "None" else v
                    else:
                        plm_analysis_ws.Cells(last_sample_count, 1).Value = "bl"
                        plm_analysis_ws.Cells(last_sample_count, 2).Value = "1"
                        plm_analysis_ws.Cells(last_sample_count, 3).Value = ""
                        plm_analysis_ws.Cells(last_sample_count, 5).Value = ""
                        plm_analysis_ws.Cells(last_sample_count, 22).Value = "100"
                        plm_analysis_ws.Cells(last_sample_count, 23).Value = "NAD"
                        plm_analysis_ws.Cells(last_sample_count, 24).Value = "50"
                        plm_analysis_ws.Cells(last_sample_count, 25).Value = "NAD"
                        plm_analysis_ws.Cells(last_sample_count, 26).Value = "50"
                        plm_analysis_ws.Cells(last_sample_count, 27).Value = "NAD"
                        plm_analysis_ws.Cells(last_sample_count, 28).Value = "50"
                        plm_analysis_ws.Cells(last_sample_count, 29).Value = "NAD"
                        plm_analysis_ws.Cells(last_sample_count, 30).Value = "50"
                else:
                    tem_analysis_ws = wb.Worksheets("TEM_Calculation")
                    count = 6
                    last_sample_count = 0
                    while True:
                        value = tem_analysis_ws.Range(f"B{count}").Value
                        if value is not None and str(value).strip() != "":
                            count += 1
                        else:
                            last_sample_count = count
                            break

                    if sample.get("whole_duplicate"):
                        for col, text_key in tem_columns_dict.items():
                            v = sample["whole_duplicate"].get(text_key, "")
                            tem_analysis_ws.Cells(last_sample_count, col).Value = "" if v == "None" else v
                    else:
                        tem_analysis_ws.Cells(last_sample_count, 1).Value = "bl"
                        tem_analysis_ws.Cells(last_sample_count, 2).Value = "1"
                        tem_analysis_ws.Cells(last_sample_count, 16).Value = "D8"
                        tem_analysis_ws.Cells(last_sample_count, 17).Value = "E8"

                wb.Save()
                processed += 1

            except Exception as e:
                errors += 1
                logging.info(f"  ✗ Ошибка обработки файла: {e}")
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

    logging.info(f"done processed={processed} errors={errors}")
