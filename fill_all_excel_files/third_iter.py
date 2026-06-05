import sys
import json
import logging
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from pathlib import Path

import win32com.client as win32


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

JSON_SOURCE_NAME = "total_data.json"

OPTIONS = ["PLM_REP", "PLM_DUP", "NOB_REP", "NOB_DUP"]

columns_dict = {
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
                ws = wb.Worksheets("SampleAnalyses")

                count = 8
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
                        v = sample["whole_duplicate"].get(text_key, "None")
                        ws.Cells(last_sample_count, col).Value = "" if v == "None" else v
                else:
                    ws.Cells(last_sample_count, 1).Value = "bl"
                    ws.Cells(last_sample_count, 2).Value = "1"
                    ws.Cells(last_sample_count, 3).Value = ""
                    ws.Cells(last_sample_count, 5).Value = ""
                    ws.Cells(last_sample_count, 22).Value = "100"
                    ws.Cells(last_sample_count, 23).Value = "NAD"
                    ws.Cells(last_sample_count, 24).Value = "50"
                    ws.Cells(last_sample_count, 25).Value = "NAD"
                    ws.Cells(last_sample_count, 26).Value = "50"
                    ws.Cells(last_sample_count, 27).Value = "NAD"
                    ws.Cells(last_sample_count, 28).Value = "50"
                    ws.Cells(last_sample_count, 29).Value = "NAD"
                    ws.Cells(last_sample_count, 30).Value = "50"

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

# pyinstaller --onefile --windowed --name="Add_Replicates_0.01" --add-data "total_data.json;." gui_app.py