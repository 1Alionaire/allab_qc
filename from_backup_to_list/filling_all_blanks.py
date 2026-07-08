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
import traceback
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

from collections import defaultdict

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

def unique_by_file_name(data):
    seen_files = set()
    result = []

    for item in data:
        file_name = item.get("file_name")

        if file_name in seen_files:
            continue

        seen_files.add(file_name)
        result.append(item)

    return result

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
            title="Выберите FILE NAMES",
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
            title="Выберите TEM BLANKS Excel-файл",
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
            all_files_data = json.load(f)


        with Path(self.file_with_weights).open("r", encoding="utf-8") as f:
            blanks_tem_data = json.load(f)

        if not all_files_data:
            self.update_status("Файлы не найдены")
            messagebox.showerror("Ошибка", "Excel-файлы не найдены в указанной папке")
            self.btn_run.config(state=NORMAL)
            self.btn_select_file_weights.config(state=NORMAL)
            return

        total_files = len(all_files_data)
        self.log_message(f"Найдено файлов: {total_files}")

        print(len(all_files_data))

        unique_files_data = unique_by_file_name(all_files_data)

        total_files = len(unique_files_data)
        self.log_message(f"Найдено файлов: {total_files}")
        pythoncom.CoInitialize()

        excel = None

        try:
            excel = win32.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            for file_index, element in enumerate(unique_files_data):
                progress_value = (file_index + 1) / total_files * 100
                self.progress["value"] = progress_value

                self.update_status(f"Обработка: {element['file_name']}")
                self.log_message("*" * 50)
                self.log_message(f"[{file_index + 1}/{total_files}] {element['file_name']}")
                
                wb = None

                try:
                    wb = open_with_retry(excel, element['file_name'])

                    if not sheet_exists(wb, "SampleAnalyses"):
                        self.log_message(f" {element['file_name']} - Лист SampleAnalyses не найден")
                        continue

                    try:
                        select_data = blanks_tem_data[element['project']]
                    except:
                        self.log_message(f" {element['file_name']} - No Blank")
                        continue

                    tem_ws = wb.Worksheets("TEM_Calculation")
                    weight_ws = wb.Worksheets("NOB_Calculation")
                    residue = None
                    last_weight_row = weight_ws.Cells(weight_ws.Rows.Count, 1).End(xlUp).Row
                    for i in range(last_weight_row, 6, -1):
                        if str(weight_ws.Range(f'A{i}').Value).lower() == 'bl':
                            residue = weight_ws.Range(f'N{i}').Value
                            break

                    if residue is None:
                        residue = 6.78

                    last_tem_row = tem_ws.Cells(tem_ws.Rows.Count, 1).End(xlUp).Row

                    find = False
                    for i in range(last_tem_row, 5, -1 ):
                        if str(tem_ws.Range(f'A{i}').Value).lower() != 'none':
                            if str(tem_ws.Range(f'A{i}').Value).lower() == 'bl':
                                find = True
                                if select_data['grid_box'] != 'none':
                                    tem_ws.Range(f'O{i}').Value = select_data['grid_box']
                                if select_data['grid_id_1'] != 'none':
                                    tem_ws.Range(f'P{i}').Value = str(select_data['grid_id_1']).upper()
                                if select_data['grid_id_2'] != 'none':
                                    tem_ws.Range(f'Q{i}').Value = str(select_data['grid_id_2']).upper()
                                tem_ws.Range(f'B{i}').Value = '1'
                                tem_ws.Range(f'E{i}').Value = residue
                                tem_ws.Range(f'F{i}').Value = 'NAD'
                                tem_ws.Range(f'G{i}').Value = 'NAD'
                                tem_ws.Range(f'I{i}').Value = 'NAD'
                                tem_ws.Range(f'J{i}').Value = 'NAD'
                                tem_ws.Range(f'L{i}').Value = 'Y'
                                tem_ws.Range(f'M{i}').Value = 'Y'
                                tem_ws.Range(f'N{i}').Value = 'Y'

                        if find == False:
                            tem_ws.Range(f'A{last_tem_row}').Value = 'bl'
                            if select_data['grid_box'] != 'none':
                                tem_ws.Range(f'O{last_tem_row}').Value = select_data['grid_box']
                            if select_data['grid_id_1'] != 'none':
                                tem_ws.Range(f'P{last_tem_row}').Value = str(select_data['grid_id_1']).upper()
                            if select_data['grid_id_2'] != 'none':
                                tem_ws.Range(f'Q{last_tem_row}').Value = str(select_data['grid_id_2']).upper()

                            tem_ws.Range(f'B{last_tem_row}').Value = '1'
                            tem_ws.Range(f'E{last_tem_row}').Value = residue
                            tem_ws.Range(f'F{last_tem_row}').Value = 'NAD'
                            tem_ws.Range(f'G{last_tem_row}').Value = 'NAD'
                            tem_ws.Range(f'I{last_tem_row}').Value = 'NAD'
                            tem_ws.Range(f'J{last_tem_row}').Value = 'NAD'
                            tem_ws.Range(f'L{last_tem_row}').Value = 'Y'
                            tem_ws.Range(f'M{last_tem_row}').Value = 'Y'
                            tem_ws.Range(f'N{last_tem_row}').Value = 'Y'

                    wb.Save()

                    self.log_message("  ✓ Файл обработан и сохранен")

                except Exception as e:
                    error_text = traceback.format_exc()

                    self.log_message(f"  ✗ Ошибка обработки файла: {e}")
                    self.log_message(error_text)

                    print(error_text)

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