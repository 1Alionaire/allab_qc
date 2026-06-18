import json
import math
import random
import copy
import os
import logging
import sys
from openpyxl import load_workbook
import win32com.client as win32
from pathlib import Path

import threading
from pathlib import Path
from tkinter import Tk, Button, Label, Text, filedialog, messagebox, END, DISABLED, NORMAL
from tkinter.ttk import Progressbar
from openpyxl import Workbook, load_workbook
import pandas as pd

import json
from collections import Counter
import random
import pythoncom

def has_duplicates(nums):
    return len(nums) != len(set(nums))

def sheet_exists(wb, sheet_name):
    for ws in wb.Worksheets:
        if ws.Name == sheet_name:
            return True
    return False

def find_project_in_sample(inp_sample_id):
    final_result = ''
    count = 2
    for i in inp_sample_id:
        if count == 0:
            return final_result[:-1]
        final_result += i
        if i == '-':
            count -= 1

def normalize_value(input_value):
    inp_str = str(input_value).strip()
    if '..' in inp_str:
        inp_str = inp_str.replace('..', '.')
    if ',' in inp_str:
        inp_str = inp_str.replace(',', '.')
    return float(inp_str)

def checking_file_name(filename):
    str_elements = ['tem', 'nob', 'plm', 'conflict', 'mold', 'air', 'lead']
    for str_element in str_elements:
        if str_element in str(filename).strip().lower():
            return False
    return True

def random_calc_point(original_value):
    operation = random.choice(['-', '+'])
    value = random.choice([1, 1.5, 2, 2.5, 3, 3.5])
    if operation == '-':
        if (float(original_value) - value) > 0:
            return float(original_value) - value
        else:
            return 2.0
    else:
        return float(original_value) + value

class QCProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Parse PCM QC")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        self.folder_path = None
        
        # Кнопка выбора папки
        self.btn_select = Button(
            root, 
            text="Choose Folder", 
            command=self.select_folder,
            width=20,
            height=2
        )
        self.btn_select.pack(pady=10)
        
        # Метка с выбранной папкой
        self.lbl_folder = Label(root, text="Папка не выбрана", wraplength=550)
        self.lbl_folder.pack(pady=5)
        
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
    
    def select_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с Excel-файлами")
        if folder:
            self.folder_path = folder
            self.lbl_folder.config(text=f"Папка: {folder}")
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
        self.btn_select.config(state=DISABLED)
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

        results = []
        pcm_sample_dict = {}
        self.update_status("Поиск Excel-файлов...")

        all_files = []
        for ext in ("*.xlsx", "*.xlsm", "*.xls"):
            all_files.extend(Path(self.folder_path).rglob(ext))

        if not all_files:
            self.update_status("Файлы не найдены")
            messagebox.showerror("Ошибка", "Excel-файлы не найдены в указанной папке")
            self.btn_run.config(state=NORMAL)
            self.btn_select.config(state=NORMAL)
            return

        total_files = len(all_files)
        self.log_message(f"Найдено файлов: {total_files}")

        pythoncom.CoInitialize()

        excel = None

        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            for file_index, filepath in enumerate(all_files):
                progress_value = (file_index + 1) / total_files * 100
                self.progress["value"] = progress_value

                self.update_status(f"Обработка: {filepath.name}")
                self.log_message(f"[{file_index + 1}/{total_files}] {filepath.name}")

                wb = None

                if checking_file_name(filepath.name) == False:
                    self.log_message(f" Wrong file")
                    continue

                try:
                    wb = excel.Workbooks.Open(str(filepath.resolve()))

                    if not sheet_exists(wb, "Count-Recount"):
                        self.log_message(f" {filepath.name} - Лист Count-Recount не найден")
                        continue

                    pcm_qc_sheet = wb.Worksheets("Count-Recount")

                    project_info = pcm_qc_sheet.Range("B2").Value
                    date_analyzed = pcm_qc_sheet.Range("L3").Value
                    analyst = str(pcm_qc_sheet.Range('S3').Value).strip()

                    if str(project_info) == 'None':
                        result_sheet = wb.Worksheets("Sample")
                        project_info = find_project_in_sample(result_sheet.Range('Q5').Value)
                    
                    if str(analyst) == 'None':
                        result_sheet = wb.Worksheets("Sample")
                        analyst = str(result_sheet.Range('B37').Value).strip()
                        if str(analyst) == 'None':
                            analyst = str(result_sheet.Range('A37').Value).strip()
                    
                    self.log_message('*' * 50)
                    self.log_message(f'project_info: {project_info}')

                    client_sample_id = None
                    original_sample_value = 0
                    qc_sample_value = 0
                    last_sample_row = 0

                    for i in range(8, 46):
                        if str(pcm_qc_sheet.Range(f"B{i}").Value).strip().lower() == 'overload':
                            continue
                        if str(pcm_qc_sheet.Range(f"B{i}").Value) == 'None':
                            last_sample_row = i - 1
                            break
                        if float(pcm_qc_sheet.Range(f"B{i}").Value) <= 0.5:
                            last_sample_row = i - 1
                            break

                    self.log_message(f'last_sample_row: {last_sample_row}')
                    have_sample = False
                    
                    for i in range(8, last_sample_row):
                        if (str(pcm_qc_sheet.Range(f'F{i}').Value) != 'None'
                            and str(pcm_qc_sheet.Range(f'G{i}').Value) != 'None'
                            and str(pcm_qc_sheet.Range(f'H{i}').Value) != 'None'):
                            have_sample = True

                            client_sample_id = str(project_info) + '-' + str(i - 7)

                            if client_sample_id in pcm_sample_dict:
                                continue
                            qc_sample_value = normalize_value(pcm_qc_sheet.Range(f'F{i}').Value)
                            original_sample_value = normalize_value(pcm_qc_sheet.Range(f'B{i}').Value)

                            pcm_sample_dict[client_sample_id] = {'original_value' : original_sample_value,
                                                                        'qc_value' : qc_sample_value,
                                                                        'analyst' : analyst, 
                                                                        'date_analyzed' : str(date_analyzed) }

                    if have_sample == False:
                        self.log_message(f'have_sample: {have_sample}')
                        if (last_sample_row - 7) < 4:
                            self.log_message(f'amount_samples: {(last_sample_row - 7)}')
                            pass
                        else:
                            random_sample_row = 0
                            amount_samples = last_sample_row - 7
                            while True:
                                random_sample_row = random.randint(8, last_sample_row)
                                amount_samples = amount_samples - 1
                                if str(pcm_qc_sheet.Range(f'B{random_sample_row}').Value).strip().lower() != 'overload':
                                    break

                                if amount_samples == 0:
                                    break
                            
                            if amount_samples == 0:
                                continue
                                
                            self.log_message(f'random_sample_row: {random_sample_row}')
                            self.log_message(f'original_sample_value:' + str(pcm_qc_sheet.Range(f'B{random_sample_row}').Value))

                            client_sample_id = str(project_info) + '-' + str(random_sample_row - 7)

                            original_sample_value = normalize_value(pcm_qc_sheet.Range(f'B{random_sample_row}').Value)
                            qc_sample_new_value = random_calc_point(original_sample_value) #250724-86

                            self.log_message(f'qc_sample_new_value: {qc_sample_new_value}')

                            pcm_qc_sheet.Range(f'F{random_sample_row}').Value = qc_sample_new_value

                            pcm_sample_dict[client_sample_id] = {'original_value' : original_sample_value,
                                                                    'qc_value' : qc_sample_new_value,
                                                                    'analyst' : analyst, 
                                                                    'date_analyzed' : str(date_analyzed) }
                            wb.Save()

                    results.append(str(filepath))

                    self.log_message("  ✓ Файл обработан и сохранен")
                except Exception as e:
                    self.log_message(f"  ✗ Ошибка обработки файла: {e}")

                finally:
                    if wb is not None:
                        wb.Close(SaveChanges=False)

        finally:
            if excel is not None:
                excel.Quit()

            pythoncom.CoUninitialize()

            self.btn_run.config(state=NORMAL)
            self.btn_select.config(state=NORMAL)
        
        output_file = Path(self.folder_path) / "qc_pcm_raw_data.json"

        with open(output_file, "a", encoding="utf-8") as f:
            f.write(json.dumps(pcm_sample_dict, indent=4, ensure_ascii=False))


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

    # pyinstaller --onefile --windowed --name="QC_PCM_Collector_2.0" parsing_all_pcm_files_2.0.py
    # pyinstaller --onefile --windowed --hidden-import win32timezone --name="QC_PCM_Collector_2.0" parsing_all_pcm_files_2.0.py
    # pyinstaller --onefile --windowed --collect-all pywin32 --name="QC_PCM_Collector_2_0" parsing_all_pcm_files_2.0.py
    # pyinstaller --onefile  --hidden-import=win32timezone  --hidden-import=win32com  --hidden-import=win32com.client   --hidden-import=pythoncom  --hidden-import=pywintypes  --name="QC_PCM_Collector_v2"   parsing_all_pcm_files_2.0.py

