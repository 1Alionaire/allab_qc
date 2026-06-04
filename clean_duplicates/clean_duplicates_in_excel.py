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

import pythoncom

def has_duplicates(nums):
    return len(nums) != len(set(nums))

def sheet_exists(wb, sheet_name):
    for ws in wb.Worksheets:
        if ws.Name == sheet_name:
            return True
    return False



def find_duplicates(nums):
    counts = Counter(nums)
    # Возвращаем ключи, количество которых больше 1
    return [item for item, count in counts.items() if count > 1]

class QCProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QC Processor")
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        
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

                try:
                    wb = excel.Workbooks.Open(str(filepath.resolve()))

                    if not sheet_exists(wb, "SampleAnalyses"):
                        self.log_message("  - Лист SampleAnalyses не найден")
                        continue

                    sample_analysis_ws = wb.Worksheets("SampleAnalyses")

                    lab_id_samples = []
                    count = 8
                    last_sample_count = 0

                    while True:
                        value = sample_analysis_ws.Range(f"B{count}").Value

                        if value is not None and str(value).strip() != "":
                            lab_id_samples.append(str(value).strip())
                            count += 1
                        else:
                            last_sample_count = count
                            break

                    if not has_duplicates(lab_id_samples):
                        self.log_message("  - Дубликаты не найдены")
                        continue

                    lab_id_dups_samples = find_duplicates(lab_id_samples)

                    self.log_message(f"  Найдены дубликаты: {lab_id_dups_samples}")

                    for row in range((last_sample_count + 4), 7, -1):
                        value = sample_analysis_ws.Range(f"B{row}").Value

                        if value is None:
                            continue

                        if str(sample_analysis_ws.Range(f"A{row}").Value).lower().strip() == 'bl':
                            for col in range(1, 50):
                                sample_analysis_ws.Cells(row, col).Value = ""

                        value = str(value).strip()

                        if value in lab_id_dups_samples:
                            for col in range(1, 50):
                                sample_analysis_ws.Cells(row, col).Value = ""
                            lab_id_dups_samples.remove(value)

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

    # pyinstaller --onefile --windowed --name="Delete_all_old_dup_samples" clean_duplicates_in_excel.py