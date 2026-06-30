from pathlib import Path
import os, sys
import pythoncom
import win32com.client as win32
xlUp = -4162


def get_writable_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)   # папка рядом с .exe
    return os.path.dirname(os.path.abspath(__file__))

script_dir = Path(__file__).resolve().parent
date = None


def generate_report_excels(input_data, input_file):
    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(input_file)
        worksheet_names = [
                    wb.Worksheets(i).Name
                    for i in range(1, wb.Worksheets.Count + 1)
                ]
        
        unique_analysts = {value["analyst"] for value in input_data.values()}

        for analyst in unique_analysts:

            if analyst in worksheet_names:
                selected_analyst_ws = wb.Worksheets(analyst)
            else:
                selected_analyst_ws = wb.Worksheets.Add()
                selected_analyst_ws.Name = analyst

            selected_analyst_data = {key: value for key, value in input_data.items() if value['analyst'] == analyst }
            
            last_row = selected_analyst_ws.Cells(selected_analyst_ws.Rows.Count, 1).End(xlUp).Row
            row = last_row + 1

            for key, value in selected_analyst_data.items():
                selected_analyst_ws.Range(f'A{row}').Value = key
                selected_analyst_ws.Range(f'B{row}').Value = value['date_analyzed'][5:7] + '-' + value['date_analyzed'][8:10] + '-' + value['date_analyzed'][0:4]
                selected_analyst_ws.Range(f'c{row}').Value = value['analyst']
                selected_analyst_ws.Range(f'D{row}').Value = value['original_value']
                selected_analyst_ws.Range(f'E{row}').Value = value['qc_value']
                row += 1

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

    
