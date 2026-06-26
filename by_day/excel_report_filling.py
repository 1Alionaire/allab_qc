import json
from pathlib import Path
import math
import random
import copy
import os
import logging
import sys
from openpyxl import load_workbook
from openpyxl import Workbook
from datetime import datetime, timedelta
import logging, traceback, os, sys

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

def decor_header_total_list(ws, type):
    ws.merge_cells('A1:A2')
    ws.merge_cells('B1:B2')
    ws.merge_cells('C1:C2')
    ws.merge_cells('D1:D2')
    ws.merge_cells('E1:G1')
    ws.merge_cells('L1:L2')

    ws['A1'] = 'QC #'
    ws['B1'] = 'Date'
    ws['C1'] = 'Batch #'
    ws['D1'] = 'Sample'
    ws['E1'] = '1st Analyst original results'
    ws['E2'] = 'Initial'
    ws['F2'] = 'Asb Type'
    ws['G2'] = 'Asbestos %'


    # DUPS
    if type == 'dup':
        ws.merge_cells('K1:K2')
        ws.merge_cells('H1:J1')
        
        ws['H1'] = 'Dups results'
        ws['H2'] = 'Initial'
        ws['I2'] = 'Asb Type'
        ws['J2'] = 'Asbestos %'
        ws['K1'] = 'R'
        ws['L1'] = 'UCL'
    else:
        ws.merge_cells('M1:M2')
        ws.merge_cells('N1:N2')
        ws.merge_cells('H1:K1')

        ws['H1'] = 'Reps results'
        ws['H2'] = 'Date'
        ws['I2'] = 'Initial'
        ws['J2'] = 'Asb Type'
        ws['K2'] = 'Asbestos %'
        ws['L1'] = 'R'
        ws['M1'] = 'LCL'
        ws['N1'] = 'UCL'

def decor_header_analyst_list(ws):
    ws.merge_cells('A1:I1')
    ws.merge_cells('A2:C2')
    ws.merge_cells('D2:F2')
    ws.merge_cells('G2:G3')
    ws.merge_cells('H2:H3')
    ws.merge_cells('I2:I3')

    ws.merge_cells('K1:S1')
    ws.merge_cells('K2:M2')
    ws.merge_cells('N2:P2')
    ws.merge_cells('Q2:Q3')
    ws.merge_cells('R2:R3')
    ws.merge_cells('S2:S3')

    ws['A1'] = 'DUPLICATE'
    ws['A2'] = '1st Analyst original result'
    ws['D2'] = '1st Analyst QC result'
    ws['G2'] = 'R Value'
    ws['H2'] = 'LCL'
    ws['I2'] = 'UCL'

    ws['A3'] = 'Initial'
    ws['B3'] = 'Asbs type'
    ws['C3'] = '%'
    ws['D3'] = 'Initial'
    ws['E3'] = 'Asbs type'
    ws['F3'] = '%'

    ws['K1'] = 'REPLICATE'
    ws['K2'] = '1st Analyst original result'
    ws['N2'] = '2nd Analyst result'
    ws['Q2'] = 'R Value'
    ws['R2'] = 'LCL'
    ws['S2'] = 'UCL'

    ws['K3'] = 'Initial'
    ws['L3'] = 'Asbs type'
    ws['M3'] = '%'
    ws['N3'] = 'Initial'
    ws['O3'] = 'Asbs type'
    ws['P3'] = '%'

types_analysis = ['plm', 'nob', 'tem']

def generate_report_excels(input_data):
    for type_analysis in types_analysis:
        chosen_type_data = [element for element in input_data if type_analysis in element['type']]
        if chosen_type_data:
            wb = Workbook()
            total_dups_ws = wb.create_sheet(f"{type_analysis.upper()} DUPs")
            total_reps_ws = wb.create_sheet(f"{type_analysis.upper()} REPs")

            decor_header_total_list(total_dups_ws, 'dup')
            decor_header_total_list(total_reps_ws, 'rep')

            text_date_for_name = chosen_type_data[0]['date'][5:7] + '-' + chosen_type_data[0]['date'][8:10] + '-' + chosen_type_data[0]['date'][0:4]
            # DUPs
            row = 3
            count = 1
            type_dup_data = [element for element in chosen_type_data if 'dup' in element['type']]
            if type_dup_data:
                for record in type_dup_data:
                    text_date_for_excel = record['date'][5:7] + '-' + record['date'][8:10] + '-' + record['date'][0:4]
                    total_dups_ws[f'A{row}'] = count
                    total_dups_ws[f'B{row}'] = text_date_for_excel
                    total_dups_ws[f'C{row}'] = record['project']
                    total_dups_ws[f'D{row}'] = record['sample']
                    total_dups_ws[f'E{row}'] = record['1st_analyst_name']
                    total_dups_ws[f'F{row}'] = record['1st_analyst_asb_type']
                    total_dups_ws[f'G{row}'] = record['1st_analyst_asb_percent']
                    total_dups_ws[f'H{row}'] = record['dup_name']
                    total_dups_ws[f'I{row}'] = record['dup_asb_type']
                    total_dups_ws[f'J{row}'] = record['dup_asb_percent']
                    total_dups_ws[f'K{row}'] = calc_std_dev(record['1st_analyst_asb_percent'], record['dup_asb_percent']) 
                    total_dups_ws[f'L{row}'] = 1
                    row += 1
                    count += 1

            # REPs
            row = 3
            count = 1
            type_rep_data = [element for element in chosen_type_data if 'rep' in element['type']]
            if type_rep_data:
                for record in type_rep_data:
                    text_date_for_excel = record['date'][5:7] + '-' + record['date'][8:10] + '-' + record['date'][0:4]
                    total_reps_ws[f'A{row}'] = count
                    total_reps_ws[f'B{row}'] = text_date_for_excel
                    total_reps_ws[f'C{row}'] = record['project']
                    total_reps_ws[f'M{row}'] = -1
                    total_reps_ws[f'N{row}'] = 1
                    if '-1000' not in record['lab id']:
                        new_date = str(datetime.strptime(text_date_for_excel, "%m-%d-%Y").date() + timedelta(days=1))
                        total_reps_ws[f'D{row}'] = record['sample']
                        total_reps_ws[f'E{row}'] = record['1st_analyst_name']
                        total_reps_ws[f'F{row}'] = record['1st_analyst_asb_type']
                        total_reps_ws[f'G{row}'] = record['1st_analyst_asb_percent']
                        total_reps_ws[f'H{row}'] = new_date[5:7] + '-' + new_date[8:10] + '-' + new_date[0:4]
                        total_reps_ws[f'I{row}'] = record['dup_name']
                        total_reps_ws[f'J{row}'] = record['dup_asb_type']
                        total_reps_ws[f'K{row}'] = record['dup_asb_percent'] 
                        total_reps_ws[f'L{row}'] = calc_std_dev(record['1st_analyst_asb_percent'], record['dup_asb_percent'])
                    else:
                        total_reps_ws[f'D{row}'] = 'BL'
                        total_reps_ws[f'E{row}'] = record['1st_analyst_name']
                        total_reps_ws[f'F{row}'] = 'NAD'
                        total_reps_ws[f'G{row}'] = "NAD"
                        total_reps_ws[f'H{row}'] = text_date_for_excel #
                        total_reps_ws[f'I{row}'] = record['dup_name'] #
                        total_reps_ws[f'J{row}'] = 'NAD'
                        total_reps_ws[f'K{row}'] = 'NAD'
                        total_reps_ws[f'L{row}'] = 0
                    row += 1
                    count += 1
            
            # Per analyst:
            unique_analysts = {item["1st_analyst_name"] for item in chosen_type_data}
            for analyst in unique_analysts:
                selected_analyst_ws = wb.create_sheet(f"{analyst}")

                decor_header_analyst_list(selected_analyst_ws)
                selected_analyst_data = [element for element in chosen_type_data if element['1st_analyst_name'] == analyst ] 
                
                if selected_analyst_data:
                    dup_row = 4
                    rep_row = 4
                    for record in selected_analyst_data:
                        if 'dup' in record['type']:
                            selected_analyst_ws[f'A{dup_row}'] = record['1st_analyst_name']
                            selected_analyst_ws[f'B{dup_row}'] = record['1st_analyst_asb_type']
                            selected_analyst_ws[f'C{dup_row}'] = record['1st_analyst_asb_percent']
                            selected_analyst_ws[f'D{dup_row}'] = record['dup_name']
                            selected_analyst_ws[f'E{dup_row}'] = record['dup_asb_type']
                            selected_analyst_ws[f'F{dup_row}'] = record['dup_asb_percent']
                            selected_analyst_ws[f'G{dup_row}'] = calc_std_dev(record['1st_analyst_asb_percent'], record['dup_asb_percent'])
                            selected_analyst_ws[f'H{dup_row}'] = -1
                            selected_analyst_ws[f'I{dup_row}'] = 1
                            dup_row += 1
                        else:
                            selected_analyst_ws[f'K{rep_row}'] = record['1st_analyst_name']
                            selected_analyst_ws[f'R{rep_row}'] = -1
                            selected_analyst_ws[f'S{rep_row}'] = 1

                            if '-1000' not in record['lab id']:
                                selected_analyst_ws[f'L{rep_row}'] = record['1st_analyst_asb_type'][:4] 
                                selected_analyst_ws[f'M{rep_row}'] = record['1st_analyst_asb_percent'] #
                                selected_analyst_ws[f'N{rep_row}'] = record['dup_name']
                                selected_analyst_ws[f'O{rep_row}'] = record['dup_asb_type'][:4]
                                selected_analyst_ws[f'P{rep_row}'] = record['dup_asb_percent']
                                selected_analyst_ws[f'Q{rep_row}'] = calc_std_dev(record['1st_analyst_asb_percent'], record['dup_asb_percent'])
                            else:
                                selected_analyst_ws[f'L{rep_row}'] = 'NAD'
                                selected_analyst_ws[f'M{rep_row}'] = 'NAD'
                                selected_analyst_ws[f'N{rep_row}'] = replicate_analyst(record['1st_analyst_name'])
                                selected_analyst_ws[f'O{rep_row}'] = 'NAD'
                                selected_analyst_ws[f'P{rep_row}'] = 'NAD'
                                selected_analyst_ws[f'Q{rep_row}'] = 0
                            rep_row += 1

            out_path = os.path.join(get_writable_dir(),  f"{type_analysis.upper()}_{text_date_for_name}.xlsx")
            wb.save(out_path)
            wb.close()


# if __name__ == '__main__':
#     file = script_dir / 'qc_result.json'

#     with file.open("r", encoding="utf-8") as f:
#         duplicates_data = json.load(f)
    
#     generate_report_excels(duplicates_data)

    
