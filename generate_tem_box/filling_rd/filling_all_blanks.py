import json
import math
import random
import traceback
import os
import logging
import sys
from openpyxl import load_workbook
import win32com.client as win32
from pathlib import Path
from collections import defaultdict
import threading
from pathlib import Path
from tkinter import Tk, Button, Label, Text, filedialog, messagebox, END, DISABLED, NORMAL
from tkinter.ttk import Progressbar
from openpyxl import Workbook, load_workbook
import pandas as pd
from datetime import date
import json
from collections import Counter
import time
import pythoncom

import time
import pythoncom
xlUp = -4162

def check_cell_value(input_value):
    if input_value is None:
        return False
    else:
        str_value = str(input_value).strip()
        if str_value == 'None' or str_value == '' or str_value == 'none':
            return False
        else:
            return True

def open_with_retry(excel, filepath, attempts=3, delay=2.0):
    """
    Открывает книгу. На транзиентных сбоях (None или COM-ошибка) ждёт и пробует снова.
    Возвращает workbook или поднимает RuntimeError после всех попыток.
    """
    last_err = None
    for i in range(1, attempts + 1):
        try:
            wb = excel.Workbooks.Open(filepath, 0, False)
            if wb is not None:
                if i > 1:
                    print(f'  ↻ Открылся со {i}-й попытки')
                return wb
            last_err = 'Open вернул None'
        except pythoncom.com_error as e:
            last_err = f'COM error: {e}'

        if i < attempts:
            time.sleep(delay * i)  # экспоненциально: 2, 4, 6 секунд

    raise RuntimeError(f'Не открылся за {attempts} попыток: {last_err}')

def has_duplicates(nums):
    return len(nums) != len(set(nums))

def sheet_exists(wb, sheet_name):
    for ws in wb.Worksheets:
        if ws.Name == sheet_name:
            return True
    return False

def get_random_weights(input_weights_data, input_residue):
    if len(input_weights_data) > 0:
        try:
            float_residue = float(input_residue)
        except:
            random_index = random.randint(0, len(input_weights_data) - 1)
            return input_weights_data.pop(random_index)
        
        interval_residue_start = float_residue - 10
        interval_residue_end = float_residue + 10

        for weights_index in range(len(input_weights_data)):
            if 0 < input_weights_data[weights_index]['percent_residue'] < 100:
                if interval_residue_start < input_weights_data[weights_index]['percent_residue'] < interval_residue_end:
                    return input_weights_data.pop(weights_index) 
                else:
                    continue
        
        random_index = random.randint(0, len(input_weights_data) - 1)
        return input_weights_data.pop(random_index)
    else:
        return
    
def find_duplicates(nums):
    counts = Counter(nums)
    # Возвращаем ключи, количество которых больше 1
    return [item for item, count in counts.items() if count > 1]

class QCProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QC Processor")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        self.file_with_file_list_path = None
        self.folder_path = None
            
        # Кнопка выбора одного файла
        self.btn_select_file_list_files = Button(
            root,
            text="Choose File with list files",
            command=self.select_json_file_with_file_list,
            width=20,
            height=2
        )
        self.btn_select_file_list_files.pack(pady=10)

        # Метка с выбранным файлом
        self.lbl_file = Label(root, text="Файл не выбран", wraplength=550)
        self.lbl_file.pack(pady=5)

        # Кнопка запуска
        self.btn_run = Button(
            root, 
            text="Запустить", 
            command=self.run_processing,
            width=20,
            height=2,
            state=DISABLED
        )
        self.btn_run.pack(pady=10)
        
        # Прогресс-бар
        self.progress = Progressbar(root, length=550, mode='determinate')
        self.progress.pack(pady=10)
        
        # Текущий статус
        self.lbl_status = Label(root, text="", font=("Arial", 10))
        self.lbl_status.pack(pady=5)
        
        # Лог
        self.log = Text(root, height=10, width=70, state=DISABLED)
        self.log.pack(pady=10)
    
    def select_json_file_with_file_list(self):
        file = filedialog.askopenfilename(
            title="Выберите FILE NAMES",
            filetypes=[
                    ("JSON", ".json")
            ]
        )
        if file:
            self.file_with_file_list_path = file
            self.lbl_file.config(text=f"Файл: {file}")
            self.btn_run.config(state=NORMAL)
    
    def log_message(self, message):
        """Добавляет сообщение в лог"""
        self.log.config(state=NORMAL)
        self.log.insert(END, message + "\n")
        self.log.see(END)
        self.log.config(state=DISABLED)
    
    def update_status(self, message):
        """Обновляет статус"""
        self.lbl_status.config(text=message)
        self.root.update()
    
    def run_processing(self):
        """Запускает обработку в отдельном потоке"""
        self.btn_run.config(state=DISABLED)
        # self.btn_select_file_weights.config(state=DISABLED)
        self.progress['value'] = 0
        
        # Очищаем лог
        self.log.config(state=NORMAL)
        self.log.delete(1.0, END)
        self.log.config(state=DISABLED)
        
        # Запускаем в отдельном потоке, чтобы GUI не зависал
        thread = threading.Thread(target=self.process_files)
        thread.start()
    
    def process_files(self):
        """Основная логика обработки файлов"""
        weights_data = None
        results = []
        samples_by_file = defaultdict(list)
        wrong_files = []
        self.update_status("Поиск Excel-файлов...")

        all_files_json = Path(self.file_with_file_list_path)

        with all_files_json.open("r", encoding="utf-8") as f:
            all_files = json.load(f)

        for sample in all_files:
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

        if not all_files:
            self.update_status("Файлы не найдены")
            messagebox.showerror("Ошибка", "Excel-файлы не найдены в указанной папке")
            self.btn_run.config(state=NORMAL)
            # self.btn_select_file_weights.config(state=NORMAL)
            return

        total_files = len(samples_by_file)
        self.log_message(f"Найдено файлов: {total_files}")

        pythoncom.CoInitialize()

        excel = None

        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            count = 0
            for file_path, samples in samples_by_file.items():
                wb = None
                plm_analysis_ws = None
                tem_analysis_ws = None
                weight_ws = None

                count += 1
                progress_value = count / total_files * 100
                self.progress["value"] = progress_value

                self.update_status(f"Обработка: {file_path}")
                self.log_message("*" * 50)
                self.log_message(f"[{count}/{total_files}] {file_path}")
                
                wb = None

                try:
                    wb = open_with_retry(excel, file_path)

                    if not sheet_exists(wb, "SampleAnalyses"):
                        self.log_message(f" {file_path} - Лист SampleAnalyses не найден")
                        continue
                    
                    tem_ws = wb.Worksheets("TEM_Calculation")
                    last_tem_row = tem_ws.Cells(tem_ws.Rows.Count, 1).End(xlUp).Row
                    print(last_tem_row)

                    for sample in samples:
                        for i in range(last_tem_row, 5, -1 ):
                            raw_sample_number = str(tem_ws.Range(f'A{i}').Value)
                            if check_cell_value(raw_sample_number): 
                                if raw_sample_number[0] in ['R', 'D'] and raw_sample_number[-2:] in ['KK', 'AB', 'VC', 'OV']:
                                    pure_sample_number = raw_sample_number[1:-2]
                                    if pure_sample_number == sample['sample']:
                                        tem_ws.Range(f'O{i}').Value = sample['Box Number']
                                        tem_ws.Range(f'P{i}').Value = str(sample['Grid_1']).upper()
                                        tem_ws.Range(f'Q{i}').Value = str(sample['Grid_2']).upper()

                    wb.Save()

                    self.log_message("  ✓ Файл обработан и сохранен")

                except Exception as e:
                    error_text = traceback.format_exc()

                    self.log_message(f"  ✗ Ошибка обработки файла: {e}")
                    self.log_message(error_text)

                    wrong_files.append(error_text)

                    continue

                finally:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
                        time.sleep(0.5) 

        finally:
            if excel is not None:
                excel.Quit()

            pythoncom.CoUninitialize()

            self.btn_run.config(state=NORMAL)

        with open("wrong_files.txt", "w", encoding="utf-8") as file:
            file.writelines(f"{item}\n" for item in wrong_files)

        if not results:
            self.update_status("Готово, но изменений не было")
            messagebox.showwarning("Внимание", "Файлы с дубликатами не найдены")
            return

        self.progress["value"] = 100
        self.update_status("Готово!")
        self.log_message(f"\n✓ Обработано файлов: {len(results)}")

        messagebox.showinfo(
            "Готово!",
            f"Обработано файлов: {len(results)}"
        )

if __name__ == "__main__":
    root = Tk()
    app = QCProcessorApp(root)
    root.mainloop()

    # pyinstaller --onefile --hidden-import babel.numbers --hidden-import babel.dates --collect-all babel --name="Filling all files with samples" filling_weights_all_files.py