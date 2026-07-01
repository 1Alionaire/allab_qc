
from pathlib import Path
import os
import logging
import sys
from openpyxl import load_workbook
from openpyxl import Workbook
from datetime import datetime, timedelta
import logging, os, sys
import pythoncom
import win32com.client as win32
import os
import sys

def get_file_path(filename):
    """ЗАПИСЬ/ЧТЕНИЕ: файлы рядом с .exe (данные, конфиги, отчёты)"""
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, filename)

xlUp = -4162

def get_writable_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)   # папка рядом с .exe
    return os.path.dirname(os.path.abspath(__file__))


script_dir = Path(__file__).resolve().parent
date = None

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

types_analysis = ['plm', 'nob', 'tem']

def generate_report_excels(input_data, files):
    for type_analysis in types_analysis:
        chosen_type_data = [element for element in input_data if type_analysis in element['type']]
        if chosen_type_data:
            pythoncom.CoInitialize()
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = None
            try:
                wb = excel.Workbooks.Open(files.get(type_analysis))

                worksheet_names = [
                    wb.Worksheets(i).Name
                    for i in range(1, wb.Worksheets.Count + 1)
                ]

                print(worksheet_names)
                total_dups_name = f"{type_analysis.upper()} DUPs"
                total_reps_name = f"{type_analysis.upper()} REPs"

                total_dups_ws = wb.Worksheets(total_dups_name)
                total_reps_ws = wb.Worksheets(total_reps_name)

                dups_last_row = total_dups_ws.Cells(total_dups_ws.Rows.Count, 8).End(xlUp).Row
                
                text_date_for_name = chosen_type_data[0]['date'][5:7] + '-' + chosen_type_data[0]['date'][8:10] + '-' + chosen_type_data[0]['date'][0:4]
                # DUPs
                row = dups_last_row + 1
                type_dup_data = [element for element in chosen_type_data if 'dup' in element['type']]

                if type_dup_data:
                    for record in type_dup_data:
                        text_date_for_excel = record['date'][5:7] + '-' + record['date'][8:10] + '-' + record['date'][0:4]
                        total_dups_ws.Range(f'B{row}').Value = text_date_for_excel
                        total_dups_ws.Range(f'C{row}').Value = record['project']
                        total_dups_ws.Range(f'D{row}').Value = record['sample']
                        total_dups_ws.Range(f'E{row}').Value = record['1st_analyst_name']
                        total_dups_ws.Range(f'F{row}').Value = record['1st_analyst_asb_type']
                        total_dups_ws.Range(f'G{row}').Value = record['1st_analyst_asb_percent']
                        total_dups_ws.Range(f'H{row}').Value = record['dup_name']
                        total_dups_ws.Range(f'I{row}').Value = record['dup_asb_type']
                        total_dups_ws.Range(f'J{row}').Value = record['dup_asb_percent']
                        total_dups_ws.Range(f'K{row}').Value = calc_std_dev(record['1st_analyst_asb_percent'], record['dup_asb_percent']) 
                        total_dups_ws.Range(f'L{row}').Value = 1
                        row += 1

                reps_last_row = total_reps_ws.Cells(total_reps_ws.Rows.Count, 8).End(xlUp).Row
                # REPs
                row = reps_last_row + 1
                type_rep_data = [element for element in chosen_type_data if 'rep' in element['type']]

                if type_rep_data:
                    for record in type_rep_data:
                        text_date_for_excel = record['date'][5:7] + '-' + record['date'][8:10] + '-' + record['date'][0:4]
                        total_reps_ws.Range(f'B{row}').Value = text_date_for_excel
                        total_reps_ws.Range(f'C{row}').Value = record['project']
                        total_reps_ws.Range(f'M{row}').Value = -1
                        total_reps_ws.Range(f'N{row}').Value = 1
                        if '-1000' not in record['lab id']:
                            new_date = str(datetime.strptime(text_date_for_excel, "%m-%d-%Y").date() + timedelta(days=1))
                            total_reps_ws.Range(f'D{row}').Value = record['sample']
                            total_reps_ws.Range(f'E{row}').Value = record['1st_analyst_name']
                            total_reps_ws.Range(f'F{row}').Value = record['1st_analyst_asb_type']
                            total_reps_ws.Range(f'G{row}').Value = record['1st_analyst_asb_percent']
                            total_reps_ws.Range(f'H{row}').Value = new_date[5:7] + '-' + new_date[8:10] + '-' + new_date[0:4]
                            total_reps_ws.Range(f'I{row}').Value = record['dup_name']
                            total_reps_ws.Range(f'J{row}').Value = record['dup_asb_type']
                            total_reps_ws.Range(f'K{row}').Value = record['dup_asb_percent'] 
                            total_reps_ws.Range(f'L{row}').Value = calc_std_dev(record['1st_analyst_asb_percent'], record['dup_asb_percent'])
                        else:
                            total_reps_ws.Range(f'D{row}').Value = 'BL'
                            total_reps_ws.Range(f'E{row}').Value = record['1st_analyst_name']
                            total_reps_ws.Range(f'F{row}').Value = 'NAD'
                            total_reps_ws.Range(f'G{row}').Value = "NAD"
                            total_reps_ws.Range(f'H{row}').Value = text_date_for_excel #
                            total_reps_ws.Range(f'I{row}').Value = record['dup_name'] #
                            total_reps_ws.Range(f'J{row}').Value = 'NAD'
                            total_reps_ws.Range(f'K{row}').Value = 'NAD'
                            total_reps_ws.Range(f'L{row}').Value = 0
                        row += 1
                
                # Per analyst:
                unique_analysts = {item["1st_analyst_name"] for item in chosen_type_data}
                for analyst in unique_analysts:
                    
                    if analyst in worksheet_names:
                        selected_analyst_ws = wb.Worksheets(analyst)
                    else:
                        selected_analyst_ws = wb.Worksheets.Add()
                        selected_analyst_ws.Name = analyst

                    selected_analyst_data = [element for element in chosen_type_data if element['1st_analyst_name'] == analyst ] 
                    
                    if selected_analyst_data:
                        dup_row = selected_analyst_ws.Cells(selected_analyst_ws.Rows.Count, 1).End(xlUp).Row
                        rep_row = selected_analyst_ws.Cells(selected_analyst_ws.Rows.Count, 11).End(xlUp).Row
                        for record in selected_analyst_data:
                            if 'dup' in record['type']:
                                selected_analyst_ws.Range(f'A{dup_row}').Value = record['1st_analyst_name']
                                selected_analyst_ws.Range(f'B{dup_row}').Value = record['1st_analyst_asb_type']
                                selected_analyst_ws.Range(f'C{dup_row}').Value = record['1st_analyst_asb_percent']
                                selected_analyst_ws.Range(f'D{dup_row}').Value = record['dup_name']
                                selected_analyst_ws.Range(f'E{dup_row}').Value = record['dup_asb_type']
                                selected_analyst_ws.Range(f'F{dup_row}').Value = record['dup_asb_percent']
                                selected_analyst_ws.Range(f'G{dup_row}').Value = calc_std_dev(record['1st_analyst_asb_percent'], record['dup_asb_percent'])
                                selected_analyst_ws.Range(f'H{dup_row}').Value = -1
                                selected_analyst_ws.Range(f'I{dup_row}').Value = 1
                                dup_row += 1
                            else:
                                selected_analyst_ws.Range(f'K{rep_row}').Value = record['1st_analyst_name']
                                selected_analyst_ws.Range(f'R{rep_row}').Value = -1
                                selected_analyst_ws.Range(f'S{rep_row}').Value = 1

                                if '-1000' not in record['lab id']:
                                    selected_analyst_ws.Range(f'L{rep_row}').Value = record['1st_analyst_asb_type'][:4] 
                                    selected_analyst_ws.Range(f'M{rep_row}').Value = record['1st_analyst_asb_percent'] #
                                    selected_analyst_ws.Range(f'N{rep_row}').Value = record['dup_name']
                                    selected_analyst_ws.Range(f'O{rep_row}').Value = record['dup_asb_type'][:4]
                                    selected_analyst_ws.Range(f'P{rep_row}').Value = record['dup_asb_percent']
                                    selected_analyst_ws.Range(f'Q{rep_row}').Value = calc_std_dev(record['1st_analyst_asb_percent'], record['dup_asb_percent'])
                                else:
                                    selected_analyst_ws.Range(f'L{rep_row}').Value = 'NAD'
                                    selected_analyst_ws.Range(f'M{rep_row}').Value = 'NAD'
                                    selected_analyst_ws.Range(f'N{rep_row}').Value = replicate_analyst(record['1st_analyst_name'])
                                    selected_analyst_ws.Range(f'O{rep_row}').Value = 'NAD'
                                    selected_analyst_ws.Range(f'P{rep_row}').Value = 'NAD'
                                    selected_analyst_ws.Range(f'Q{rep_row}').Value = 0
                                rep_row += 1

                wb.Save()

            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
                excel.Quit()
                pythoncom.CoUninitialize()

# if __name__ == '__main__':
#     file = script_dir / 'qc_result.json'

#     with file.open("r", encoding="utf-8") as f:
#         duplicates_data = json.load(f)
    
#     generate_report_excels(duplicates_data)

    
