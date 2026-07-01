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
from datetime import date
import json
from collections import Counter
import time
import pythoncom

import time
import pythoncom
xlUp = -4162

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
        self.file_with_weights = None
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

        # Кнопка выбора одного файла
        self.btn_select_file_weights = Button(
            root,
            text="Choose File with weights",
            command=self.select_json_file_with_weights,
            width=20,
            height=2
        )
        self.btn_select_file_weights.pack(pady=10)

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
            title="Выберите Excel-файл",
            filetypes=[
                    ("JSON", ".json")
            ]
        )
        if file:
            self.file_with_file_list_path = file
            self.lbl_file.config(text=f"Файл: {file}")
            self.btn_run.config(state=NORMAL)

    def select_json_file_with_weights(self):
        file = filedialog.askopenfilename(
            title="Выберите Excel-файл",
            filetypes=[
                    ("JSON", ".json")
            ]
        )
        if file:
            self.file_with_weights = file
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
        self.btn_select_file_weights.config(state=DISABLED)
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

        self.update_status("Поиск Excel-файлов...")

        all_files_json = Path(self.file_with_file_list_path)
        with all_files_json.open("r", encoding="utf-8") as f:
            all_files = json.load(f)   
        print(all_files)

        with Path(self.file_with_weights).open("r", encoding="utf-8") as f:
            weights_data = json.load(f)

        if not all_files:
            self.update_status("Файлы не найдены")
            messagebox.showerror("Ошибка", "Excel-файлы не найдены в указанной папке")
            self.btn_run.config(state=NORMAL)
            self.btn_select_file_weights.config(state=NORMAL)
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

                self.update_status(f"Обработка: {filepath}")
                self.log_message("*" * 50)
                self.log_message(f"[{file_index + 1}/{total_files}] {filepath}")
                
                wb = None

                try:
                    wb = open_with_retry(excel, filepath)

                    if not sheet_exists(wb, "SampleAnalyses"):
                        self.log_message(f" {filepath} - Лист SampleAnalyses не найден")
                        continue

                    sample_analysis_ws = wb.Worksheets("SampleAnalyses")
                    report_ws = wb.Worksheets("PLM_TEM_Report")
                    weight_ws = wb.Worksheets("NOB_Calculation")

                    if (sample_analysis_ws.Range("A8").Value is None and str(sample_analysis_ws.Range("A8").Value).strip() == ""
                        and sample_analysis_ws.Range("B8").Value is None and str(sample_analysis_ws.Range("B8").Value).strip() == ""):
                        self.log_message(f"{filepath} - No Samples")
                        continue

                    if str(sample_analysis_ws.Range("A8").Value) == 'None' and str(sample_analysis_ws.Range("B8").Value) == 'None':
                        self.log_message(f"{filepath} - No Samples")
                        continue
                    
                    report_row = 6
                    total_amount_samples = 0
                    while True:
                        lab_id = report_ws.Range(f"B{report_row}").Value
                        if lab_id is None:
                            total_amount_samples = report_row - 5
                            break
                        else:
                            if str(lab_id).strip == '':
                                total_amount_samples = report_row - 5
                                break
                        report_row += 1

                    sample_row = 7 + total_amount_samples

                    self.log_message(f" total_amount_samples: {total_amount_samples - 1} ")

                    duplicates_samples = []
                    while True:
                        sample_client_id = sample_analysis_ws.Range(f"A{sample_row}").Value
                        sample_lab_id = sample_analysis_ws.Range(f"B{sample_row}").Value

                        if (sample_client_id is not None and str(sample_client_id).strip() != ""
                            and sample_lab_id is not None and str(sample_lab_id).strip() != "") and (str(sample_client_id).strip() != 'None' and str(sample_lab_id).strip() != "None"):
                            if str(sample_client_id).lower() != 'bl' or str(sample_lab_id).lower() != '1':
                                if sample_analysis_ws.Range(f"AR{sample_row}").Value is not None and sample_analysis_ws.Range(f"AR{sample_row}").Value != '':
                                    if str(sample_analysis_ws.Range(f"AR{sample_row}").Value) == '198.6':
                                        print(f'Sample number: {sample_analysis_ws.Cells(sample_row, 1).Value,}')
                                        print(f'lab_id: {sample_analysis_ws.Cells(sample_row, 2).Value,}')
                                        print(f'residue: {sample_analysis_ws.Cells(sample_row, 46).Value,}')
                                        duplicates_samples.append({
                                            'sample_number':sample_analysis_ws.Cells(sample_row, 1).Value,
                                            'lab_id':sample_analysis_ws.Cells(sample_row, 2).Value,
                                            'residue': sample_analysis_ws.Cells(sample_row, 46).Value
                                        })
                            # self.log_message(f"will delete {sample_analysis_ws.Cells(sample_row, 1).Value}")
                            # for col in range(1, 50):
                                # sample_analysis_ws.Cells(sample_row, col).Value = ""
                        else:
                            break
                        sample_row += 1

                    last_weight_row = weight_ws.Cells(weight_ws.Rows.Count, 1).End(xlUp).Row
                    weight_row = last_weight_row + 1

                    for duplicate in duplicates_samples:
                        weight_item = get_random_weights(weights_data, duplicate['residue'])
                        if weight_item is not None:
                            weight_ws.Range(f'A{weight_row}').Value = duplicate['sample_number']
                            weight_ws.Range(f'B{weight_row}').Value = duplicate['lab_id']
                            weight_ws.Range(f'C{weight_row}').Value = '1'
                            weight_ws.Range(f'D{weight_row}').Value = weight_item['cruc_weight']
                            weight_ws.Range(f'E{weight_row}').Value = weight_item['cruc_with_sample_weight']
                            weight_ws.Range(f'F{weight_row}').Value = weight_item['sample_weight']
                            weight_ws.Range(f'G{weight_row}').Value = weight_item['cruc_with_sample_ash_weight']
                            weight_ws.Range(f'H{weight_row}').Value = weight_item['percent_organic']
                            weight_ws.Range(f'I{weight_row}').Value = weight_item['petri_weight']
                            weight_ws.Range(f'J{weight_row}').Value = weight_item['petri_with_sample_weight']
                            weight_ws.Range(f'K{weight_row}').Value = weight_item['petri_with_sample_weight']
                            weight_ws.Range(f'L{weight_row}').Value = 0
                            weight_ws.Range(f'M{weight_row}').Value = weight_item['percent_caco3']
                            weight_ws.Range(f'N{weight_row}').Value = weight_item['percent_residue']
                            weight_ws.Range(f'O{weight_row}').Value = '198.6'
                        else:
                            print(f'ERROR {filepath}')
                        weight_row += 1

                    wb.Save()

                    #results.append(str(filepath))

                    self.log_message("  ✓ Файл обработан и сохранен")

                except Exception as e:
                    self.log_message(f"  ✗ Ошибка обработки файла: {e}")
                    continue

                finally:
                    if wb is not None:
                        wb.Close(SaveChanges=False)
                        time.sleep(0.3) 

        finally:
            if excel is not None:
                excel.Quit()

            pythoncom.CoUninitialize()

            self.btn_run.config(state=NORMAL)
            self.btn_select_file_weights.config(state=NORMAL)

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