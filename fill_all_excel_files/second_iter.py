from pathlib import Path
import json
import logging
from openpyxl import load_workbook
import win32com.client as win32
from pathlib import Path

script_dir = Path(__file__).resolve().parent
log_file = script_dir / "app.log"

# Удаляем старые обработчики logging
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

def get_resource_path(name):
    """Путь к ресурсу — внутри .exe или рядом с .py."""
    if getattr(sys, 'frozen', False):
        return Path(sys._MEIPASS) / name   # внутри bundle
    return Path(__file__).resolve().parent / name

# Настраиваем логирование в файл
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(str(log_file), encoding="utf-8")
    ]
)

columns_dict = {
    1 : "Client ID",
    2 : "Lab ID",
    3 :  "Layer",
    4 :  "Color",
    5 :  "Texture",
    6 :  "Homogeneity",
    7:  "Morphology",
    8 :  "RI II Type 1",
    9 :  "RI II Type 2",
    10: "RI ┴ Type 1",
    11:  "RI ┴ Type 2",
    12:  "Sign of \nElongation Type 1",
    13:  "Sign of \nElongation Type 2",
    14:  "Extinction \nAngle Type 1",
    15:  "Extinction \nAngle Type 2",
    16:  "Pleochroism /\nColor Type 1",
    17:  "Pleochroism /\nColor Type 2",
    18:  "Birefringence Type 1",
    19: "Birefringence Type 2",
    20:  "Other Fibers",
    21:  "Property",
    22:  "% Non-\nAsbestos", 
    23:  "Type 1",
    24:  "Point 1",
    25:  "Type 2",
    26:  "Point 2",
    27:  "Type 3",
    28:  "Point 3",
    29:  "Type 4",
    30:  "Point 4",
    31:  "Type 5",
    32:  "Point 5",
    33:  "Type 6",
    34:  "Point 6",
    35:  "Type 7",
    36:  "Point 7",
    37:  "Type 8",
    38:  "Point 8",
    39:  "Type Asb 1 Option",
    40:  "Percent 1 Option",
    41:  "Type Asb 2 Option",
    42:  "Percent 2 Option",
    43:  "Vermiculite", 
    44:  "Method", 
    45:  "Undesolved Materials",
    46:  "Total Residue",
}

script_dir = Path(__file__).resolve().parent

json_source_name = 'PLM_DUP.json'

json_path = get_resource_path(json_source_name)

with json_path.open("r", encoding="utf-8") as f:
    all_samples = json.load(f)

excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.AutomationSecurity = 1  # без Protected View

chosen_company = 'ABC Environmental LLC'

chosen_company_data = [value for value in all_samples if (chosen_company in value['file_name'] )]
                                                                        
logging.info(f'chosen_company_data: {chosen_company_data}')

try:
    for sample in chosen_company_data:
        logging.info(sample['file_name'])

        wb = None
        try:
            wb = excel.Workbooks.Open(str(Path(sample['file_name']).resolve()))
            sample_analysis_ws = wb.Worksheets("SampleAnalyses")

            count = 8
            last_sample_count = 0

            while True:
                value = sample_analysis_ws.Range(f"B{count}").Value

                if value is not None and str(value).strip() != "":
                    count += 1
                else:
                    last_sample_count = count
                    break

            if sample['whole_duplicate']:
                for col, text_key in columns_dict.items():
                    if sample['whole_duplicate'][text_key] != 'None':
                        sample_analysis_ws.Cells(last_sample_count, col).Value = sample['whole_duplicate'][text_key]
                    else:
                        sample_analysis_ws.Cells(last_sample_count, col).Value = ''
            else:
                sample_analysis_ws.Cells(last_sample_count, 1).Value = 'bl'
                sample_analysis_ws.Cells(last_sample_count, 2).Value = '1'
                sample_analysis_ws.Cells(last_sample_count, 3).Value = ''
                sample_analysis_ws.Cells(last_sample_count, 5).Value = ''
                sample_analysis_ws.Cells(last_sample_count, 22).Value = '100'
                sample_analysis_ws.Cells(last_sample_count, 23).Value = 'NAD'
                sample_analysis_ws.Cells(last_sample_count, 24).Value = '50'
                sample_analysis_ws.Cells(last_sample_count, 25).Value = 'NAD'
                sample_analysis_ws.Cells(last_sample_count, 26).Value = '50'
                sample_analysis_ws.Cells(last_sample_count, 27).Value = 'NAD'
                sample_analysis_ws.Cells(last_sample_count, 28).Value = '50'
                sample_analysis_ws.Cells(last_sample_count, 29).Value = 'NAD'
                sample_analysis_ws.Cells(last_sample_count, 30).Value = '50'


            wb.Save()


        except Exception as e:
            logging.info(f"  ✗ Ошибка обработки файла: {e}")

        finally:
            if wb is not None:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
finally:
    excel.Quit()

# pyinstaller --onefile --windowed --name="Add_PLM_Replicates_0.01" first_iter.py
# with file: pyinstaller --onefile --windowed --name="Add_PLM_Replicates_0.01" --add-data "PLM_REP_test_data.json;." first_iter.py
# with file: pyinstaller --onefile --windowed --name="Add_PLM_Duplicates_0.01" --add-data "PLM_DUP.json;." second_iter.py