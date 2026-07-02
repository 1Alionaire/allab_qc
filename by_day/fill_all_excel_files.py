from collections import defaultdict
from pathlib import Path
import gc
import logging
import pythoncom
import win32com.client as win32
import shutil
import tempfile
import time
import sys
import json 
import random
xlUp = -4162

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

def get_random_weights(input_weights_data, input_residue):
    if len(input_weights_data) > 0:
        try:
            float_residue = float(input_residue)
        except:
            random_index = random.randint(0, len(input_weights_data) - 1)
            return input_weights_data.pop(random_index)
        
        interval_residue_start = float_residue - 10
        interval_residue_end = float_residue + 10

        while True:
            random_index = random.randint(0, len(input_weights_data) - 1)
            if 0 < input_weights_data[random_index]['percent_residue'] < 100:
                if interval_residue_start < input_weights_data[random_index]['percent_residue'] < interval_residue_end:
                    return input_weights_data.pop(random_index) 
            if len(input_weights_data) == 0:
                break
    else:
        return

def get_resource_path(filename):
    """
    Returns the path to a file bundled inside the executable.

    When running as a normal .py file:
        returns the directory containing the script.

    When running as a PyInstaller .exe:
        returns PyInstaller's temporary resource directory.
    """
    if getattr(sys, "frozen", False):
        base_dir = Path(sys._MEIPASS)
    else:
        base_dir = Path(__file__).resolve().parent

    return base_dir / filename

def copy_with_retry(src, dst, attempts=5, delay=2):
    """Копирование с повторами — на случай если файл на Synology занят."""
    for i in range(attempts):
        try:
            shutil.copy2(src, dst)
            return
        except (PermissionError, OSError) as e:
            if i == attempts - 1:
                raise
            logging.warning(
                "Копирование занято, повтор %s/%s: %s",
                i + 1, attempts, e
            )
            time.sleep(delay)

def fill_all_excel_files(inp_data):
    samples_by_file = defaultdict(list)

    weight_config_path = get_resource_path("correct_weight.json")
    with weight_config_path.open("r", encoding="utf-8") as file:
        weight_data = json.load(file)

    for sample in inp_data:
        fname = sample.get("file_name", "")

        if not fname:
            logging.warning("Sample без file_name: %s", sample)
            continue

        file_path = Path(fname).resolve()
        print(file_path)
        if not file_path.exists():
            logging.error("Файл не существует: %s", file_path)
            continue

        if file_path.name.startswith("~$"):
            logging.warning("Пропущен lock-файл: %s", file_path)
            continue

        if file_path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            logging.warning("Не Excel-файл: %s", file_path)
            continue

        samples_by_file[file_path].append(sample)

    pythoncom.CoInitialize()

    excel = None
    processed = 0
    errors = 0

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.EnableEvents = False

        # Отключаем макросы
        excel.AutomationSecurity = 3

        for file_path, samples in samples_by_file.items():
            wb = None
            plm_analysis_ws = None
            tem_analysis_ws = None
            weight_ws = None

            # Временная локальная папка — Excel не касается сети
            local_dir = Path(tempfile.mkdtemp())
            local_path = local_dir / file_path.name

            weights_samples = []
            try:
                logging.info(
                    "Opening: %s | samples: %s",
                    file_path,
                    len(samples)
                )

                # Сеть → локально
                copy_with_retry(file_path, local_path)

                wb = excel.Workbooks.Open(
                    str(local_path),
                    UpdateLinks=0,
                    ReadOnly=False,
                    IgnoreReadOnlyRecommended=True,
                    Notify=False,
                    AddToMru=False
                )

                logging.info(
                    "Workbook.Name=%s | FullName=%s | ReadOnly=%s",
                    wb.Name,
                    wb.FullName,
                    wb.ReadOnly
                )

                if wb.ReadOnly:
                    raise PermissionError(
                        f"Файл открыт только для чтения: {local_path}"
                    )

                for sample in samples:
                    type_analysis = str(
                        sample.get("type", "")
                    ).lower()

                    logging.info(
                        "Project=%s | type=%s",
                        sample.get("project", ""),
                        type_analysis
                    )

                    if "nob" in type_analysis or "plm" in type_analysis:
                        plm_analysis_ws = wb.Worksheets(
                            "SampleAnalyses"
                        )

                        row_number = 8

                        while True:
                            value = plm_analysis_ws.Range(
                                f"B{row_number}"
                            ).Value

                            if value is None or str(value).strip() == "":
                                break

                            row_number += 1

                        duplicate = sample.get("whole_duplicate")

                        if duplicate:
                            for col, text_key in plm_columns_dict.items():
                                value = duplicate.get(text_key)

                                if value in (None, "None"):
                                    value = ""

                                plm_analysis_ws.Cells(
                                    row_number,
                                    col
                                ).Value = value

                        else:
                            plm_analysis_ws.Cells(
                                row_number, 1
                            ).Value = "bl"

                            plm_analysis_ws.Cells(
                                row_number, 2
                            ).Value = "1"

                            plm_analysis_ws.Cells(
                                row_number, 3
                            ).Value = ""

                            plm_analysis_ws.Cells(
                                row_number, 5
                            ).Value = ""

                            plm_analysis_ws.Cells(
                                row_number, 22
                            ).Value = 100

                            for type_col, point_col in (
                                (23, 24),
                                (25, 26),
                                (27, 28),
                                (29, 30),
                            ):
                                plm_analysis_ws.Cells(
                                    row_number,
                                    type_col
                                ).Value = "NAD"

                                plm_analysis_ws.Cells(
                                    row_number,
                                    point_col
                                ).Value = 50

                        if "nob" in type_analysis:
                            if duplicate:
                                original_sample_number = str(sample.get("sample", "")).lower()
                                # we add original sample number in case replicate and duplicate has a same original sample number. if duplicate and replicate have same sample number - only one weight
                                if original_sample_number in weights_samples:
                                    pass
                                else:
                                    weight_ws = wb.Worksheets("NOB_Calculation")
                                    weight_item = get_random_weights(weight_data, duplicate.get('Total Residue'))
                                    last_weight_row = weight_ws.Cells(weight_ws.Rows.Count, 1).End(xlUp).Row + 1
                                    
                                    if weight_item is not None:

                                        weights_samples.append(original_sample_number)

                                        weight_ws.Range(f'A{last_weight_row}').Value = original_sample_number + 'r'
                                        weight_ws.Range(f'B{last_weight_row}').Value = str(sample.get("lab id", ""))
                                        weight_ws.Range(f'C{last_weight_row}').Value = '1'
                                        weight_ws.Range(f'D{last_weight_row}').Value = weight_item['cruc_weight']
                                        weight_ws.Range(f'E{last_weight_row}').Value = weight_item['cruc_with_sample_weight']
                                        weight_ws.Range(f'F{last_weight_row}').Value = weight_item['sample_weight']
                                        weight_ws.Range(f'G{last_weight_row}').Value = weight_item['cruc_with_sample_ash_weight']
                                        weight_ws.Range(f'H{last_weight_row}').Value = weight_item['percent_organic']
                                        weight_ws.Range(f'I{last_weight_row}').Value = weight_item['petri_weight']
                                        weight_ws.Range(f'J{last_weight_row}').Value = weight_item['petri_with_sample_weight']
                                        weight_ws.Range(f'K{last_weight_row}').Value = weight_item['petri_with_sample_weight']
                                        weight_ws.Range(f'L{last_weight_row}').Value = 0
                                        weight_ws.Range(f'M{last_weight_row}').Value = weight_item['percent_caco3']
                                        weight_ws.Range(f'N{last_weight_row}').Value = weight_item['percent_residue']
                                        weight_ws.Range(f'O{last_weight_row}').Value = '198.6'

                                        plm_analysis_ws.Range(f'AT{row_number}').Value = weight_item['percent_residue']
                    else:
                        tem_analysis_ws = wb.Worksheets(
                            "TEM_Calculation"
                        )

                        row_number = 6

                        while True:
                            value = tem_analysis_ws.Range(
                                f"B{row_number}"
                            ).Value

                            if value is None or str(value).strip() == "":
                                break

                            row_number += 1

                        duplicate = sample.get("whole_duplicate")

                        if duplicate:
                            for col, text_key in tem_columns_dict.items():
                                value = duplicate.get(text_key)

                                if value in (None, "None"):
                                    value = ""

                                tem_analysis_ws.Cells(
                                    row_number,
                                    col
                                ).Value = value

                        else:
                            tem_analysis_ws.Cells(
                                row_number, 1
                            ).Value = "bl"

                            tem_analysis_ws.Cells(
                                row_number, 2
                            ).Value = "1"

                            tem_analysis_ws.Cells(
                                row_number, 16
                            ).Value = "D8"

                            tem_analysis_ws.Cells(
                                row_number, 17
                            ).Value = "E8"

                    processed += 1

                

                logging.info("Saving locally: %s", local_path)
                wb.Save()
                wb.Close(SaveChanges=False)
                wb = None
                gc.collect()

                # Локально → сеть
                logging.info("Copying back to: %s", file_path)
                copy_with_retry(local_path, file_path)
                logging.info("Saved successfully: %s", file_path)

            except Exception as error:
                errors += 1

                logging.exception(
                    "Ошибка обработки файла: %s",
                    file_path
                )

                print(f"Ошибка обработки {file_path}: {error}")
                # raise убран — один битый файл не роняет весь прогон

            finally:
                plm_analysis_ws = None
                tem_analysis_ws = None

                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        logging.exception(
                            "Ошибка закрытия книги: %s",
                            local_path
                        )

                wb = None
                shutil.rmtree(local_dir, ignore_errors=True)
                gc.collect()
    finally:
        if excel is not None:
            try:
                excel.EnableEvents = True
            except Exception:
                pass

            try:
                excel.Quit()
            except Exception:
                logging.exception("Ошибка закрытия Excel")

        excel = None
        gc.collect()
        pythoncom.CoUninitialize()

    logging.info(
        "Done: processed=%s errors=%s",
        processed,
        errors
    )