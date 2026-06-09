import sys
import json
import logging
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from pathlib import Path

import win32com.client as win32
import pythoncom


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


# --- логирование в файл ---
log_file = get_external_path("app.log")
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.FileHandler(str(log_file), encoding="utf-8")],
)

JSON_SOURCE_NAME = "tem_total_data.json"

OPTIONS = ["TEM_REP", "TEM_DUP"]

columns_dict = {
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

def process_samples(option, filter_text, log_callback):
    """Основная логика. log_callback(str) — для вывода в GUI."""
    json_path = get_resource_path(JSON_SOURCE_NAME)
    with json_path.open("r", encoding="utf-8") as f:
        all_data = json.load(f)

    if option not in all_data:
        raise ValueError(f"Опция '{option}' не найдена в данных")

    samples = all_data[option]

    # фильтрация по тексту в file_name
    filter_text = (filter_text or "").strip()
    if filter_text:
        chosen = [s for s in samples if filter_text in s.get("file_name", "")]
    else:
        chosen = samples

    log_callback(f"Опция: {option}")
    log_callback(f"Фильтр: '{filter_text}'" if filter_text else "Фильтр: (нет)")
    log_callback(f"Найдено файлов: {len(chosen)}")
    logging.info(f"option={option} filter='{filter_text}' count={len(chosen)}")

    if not chosen:
        log_callback("Нет подходящих файлов. Остановлено.")
        return

    pythoncom.CoInitialize()
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1  # без Protected View

    processed = 0
    errors = 0
    try:
        for sample in chosen:
            fname = sample.get("file_name", "")
            log_callback(f"→ {fname}")
            logging.info(fname)

            wb = None
            try:
                wb = excel.Workbooks.Open(str(Path(fname).resolve()))
                ws = wb.Worksheets("TEM_Calculation")

                count = 6
                last_sample_count = 0
                while True:
                    value = ws.Range(f"B{count}").Value
                    if value is not None and str(value).strip() != "":
                        count += 1
                    else:
                        last_sample_count = count
                        break

                if sample.get("whole_duplicate"):
                    for col, text_key in columns_dict.items():
                        v = sample["whole_duplicate"].get(text_key, "")
                        ws.Cells(last_sample_count, col).Value = "" if v == "None" else v
                else:
                    ws.Cells(last_sample_count, 1).Value = "bl"
                    ws.Cells(last_sample_count, 2).Value = "1"
                    ws.Cells(last_sample_count, 16).Value = "D8"
                    ws.Cells(last_sample_count, 17).Value = "E8"

                wb.Save()
                processed += 1
                log_callback("   ✓ сохранено")

            except Exception as e:
                errors += 1
                logging.info(f"  ✗ Ошибка обработки файла: {e}")
                log_callback(f"   ✗ ошибка: {e}")
            finally:
                if wb is not None:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()

    log_callback(f"\nГотово. Обработано: {processed}, ошибок: {errors}")
    logging.info(f"done processed={processed} errors={errors}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Add Replicates / Duplicates")
        self.geometry("640x520")

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="Текст для фильтрации (по file_name):").pack(anchor="w")
        self.entry = ttk.Entry(frm, width=70)
        self.entry.pack(fill="x", pady=(2, 10))

        ttk.Label(frm, text="Тип обработки:").pack(anchor="w")
        self.option_var = tk.StringVar(value=OPTIONS[0])
        self.combo = ttk.Combobox(
            frm, textvariable=self.option_var, values=OPTIONS, state="readonly"
        )
        self.combo.pack(fill="x", pady=(2, 10))

        self.start_btn = ttk.Button(frm, text="Старт", command=self.on_start)
        self.start_btn.pack(pady=(0, 10))

        ttk.Label(frm, text="Лог:").pack(anchor="w")
        self.log_box = scrolledtext.ScrolledText(frm, height=18, state="disabled")
        self.log_box.pack(fill="both", expand=True)

    def log(self, msg):
        def _append():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, _append)

    def on_start(self):
        option = self.option_var.get()
        filter_text = self.entry.get()

        self.start_btn.configure(state="disabled")
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        def worker():
            try:
                process_samples(option, filter_text, self.log)
            except Exception as e:
                logging.exception("Фатальная ошибка")
                self.log(f"ФАТАЛЬНАЯ ОШИБКА: {e}")
                self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
            finally:
                self.after(0, lambda: self.start_btn.configure(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()

# pyinstaller --onefile --windowed --name="TEM_Add_Replicates_0.01" --add-data "tem_total_data.json;." tem_fourth_iter.py
