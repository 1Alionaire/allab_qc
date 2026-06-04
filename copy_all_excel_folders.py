import shutil
import threading
from pathlib import Path
from tkinter import Tk, Button, Label, Text, filedialog, messagebox, END, DISABLED, NORMAL


EXCEL_EXTENSIONS = {
    ".xls",
    ".xlsx",
    ".xlsm",
    ".xlsb",
    ".xlt",
    ".xltx",
    ".xltm",
}


class ExcelFolderCopierApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Folder Copier")
        self.root.geometry("700x450")
        self.root.resizable(False, False)

        self.source_folder = None
        self.destination_folder = None
        self.is_copying = False

        self.btn_source = Button(
            root,
            text="1. Выбрать папку, откуда копировать",
            width=40,
            height=2,
            command=self.select_source_folder
        )
        self.btn_source.pack(pady=15)

        self.lbl_source = Label(root, text="Источник: не выбран", wraplength=650)
        self.lbl_source.pack(pady=5)

        self.btn_destination = Button(
            root,
            text="2. Выбрать папку, куда копировать",
            width=40,
            height=2,
            command=self.select_destination_folder
        )
        self.btn_destination.pack(pady=15)

        self.lbl_destination = Label(root, text="Назначение: не выбрано", wraplength=650)
        self.lbl_destination.pack(pady=5)

        self.lbl_status = Label(root, text="", font=("Arial", 10))
        self.lbl_status.pack(pady=10)

        self.log = Text(root, height=12, width=85, state=DISABLED)
        self.log.pack(pady=10)

    def log_message(self, message):
        self.log.config(state=NORMAL)
        self.log.insert(END, message + "\n")
        self.log.see(END)
        self.log.config(state=DISABLED)

    def set_status(self, message):
        self.lbl_status.config(text=message)

    def select_source_folder(self):
        if self.is_copying:
            messagebox.showwarning("Подождите", "Копирование уже выполняется.")
            return

        folder = filedialog.askdirectory(title="Выберите папку, откуда копировать")
        if folder:
            self.source_folder = Path(folder)
            self.lbl_source.config(text=f"Источник: {self.source_folder}")
            self.try_start_copying()

    def select_destination_folder(self):
        if self.is_copying:
            messagebox.showwarning("Подождите", "Копирование уже выполняется.")
            return

        folder = filedialog.askdirectory(title="Выберите папку, куда копировать")
        if folder:
            self.destination_folder = Path(folder)
            self.lbl_destination.config(text=f"Назначение: {self.destination_folder}")
            self.try_start_copying()

    def try_start_copying(self):
        """
        Так как пользователь просил только две кнопки,
        копирование запускается автоматически после выбора обеих папок.
        """
        if self.source_folder and self.destination_folder:
            if not self.validate_folders():
                return

            answer = messagebox.askyesno(
                "Начать копирование?",
                "Обе папки выбраны.\n\nНачать копирование Excel-файлов?"
            )

            if answer:
                self.start_copying_thread()

    def is_inside_folder(self, child, parent):
        """
        Проверяет, находится ли child внутри parent.
        Работает в Python 3.8+.
        """
        try:
            child.resolve().relative_to(parent.resolve())
            return True
        except ValueError:
            return False

    def validate_folders(self):
        source = self.source_folder.resolve()
        destination = self.destination_folder.resolve()

        if source == destination:
            messagebox.showerror(
                "Ошибка",
                "Папка-источник и папка-назначение не должны быть одинаковыми."
            )
            return False

        if self.is_inside_folder(destination, source):
            messagebox.showerror(
                "Ошибка",
                "Папка назначения не должна находиться внутри папки-источника.\n\n"
                "Иначе программа может начать копировать результаты внутрь самой себя."
            )
            return False

        return True

    def start_copying_thread(self):
        self.is_copying = True
        self.btn_source.config(state=DISABLED)
        self.btn_destination.config(state=DISABLED)

        self.log.config(state=NORMAL)
        self.log.delete("1.0", END)
        self.log.config(state=DISABLED)

        thread = threading.Thread(target=self.copy_excel_files)
        thread.daemon = True
        thread.start()

    def get_unique_destination_path(self, destination_file):
        """
        Если файл уже существует, не перезаписывает его,
        а создает имя вида:
        report.xlsx
        report_copy_1.xlsx
        report_copy_2.xlsx
        """
        if not destination_file.exists():
            return destination_file

        folder = destination_file.parent
        stem = destination_file.stem
        suffix = destination_file.suffix

        counter = 1

        while True:
            new_file = folder / f"{stem}_copy_{counter}{suffix}"
            if not new_file.exists():
                return new_file
            counter += 1

    def copy_excel_files(self):
        try:
            source_root = self.source_folder.resolve()
            destination_root = self.destination_folder.resolve()

            self.root.after(0, self.set_status, "Поиск Excel-файлов...")
            self.root.after(0, self.log_message, f"Источник: {source_root}")
            self.root.after(0, self.log_message, f"Назначение: {destination_root}")
            self.root.after(0, self.log_message, "-" * 70)

            excel_files = []

            for file_path in source_root.rglob("*"):
                if not file_path.is_file():
                    continue

                if file_path.name.startswith("~$"):
                    continue

                if file_path.suffix.lower() in EXCEL_EXTENSIONS:
                    excel_files.append(file_path)

            if not excel_files:
                self.root.after(0, self.set_status, "Excel-файлы не найдены.")
                self.root.after(0, self.log_message, "Excel-файлы не найдены.")
                self.root.after(0, messagebox.showwarning, "Внимание", "Excel-файлы не найдены.")
                return

            self.root.after(0, self.log_message, f"Найдено Excel-файлов: {len(excel_files)}")
            self.root.after(0, self.log_message, "-" * 70)

            copied_count = 0
            error_count = 0

            for index, source_file in enumerate(excel_files, start=1):
                try:
                    relative_path = source_file.relative_to(source_root)
                    destination_file = destination_root / relative_path

                    destination_file.parent.mkdir(parents=True, exist_ok=True)

                    final_destination_file = self.get_unique_destination_path(destination_file)

                    shutil.copy2(source_file, final_destination_file)

                    copied_count += 1

                    self.root.after(
                        0,
                        self.log_message,
                        f"[{index}/{len(excel_files)}] ✓ {relative_path}"
                    )

                    self.root.after(
                        0,
                        self.set_status,
                        f"Копирование: {index}/{len(excel_files)}"
                    )

                except Exception as e:
                    error_count += 1
                    self.root.after(
                        0,
                        self.log_message,
                        f"[{index}/{len(excel_files)}] ✗ Ошибка: {source_file} | {e}"
                    )

            final_message = (
                f"Готово!\n\n"
                f"Скопировано файлов: {copied_count}\n"
                f"Ошибок: {error_count}"
            )

            self.root.after(0, self.set_status, "Готово!")
            self.root.after(0, self.log_message, "-" * 70)
            self.root.after(0, self.log_message, f"Скопировано файлов: {copied_count}")
            self.root.after(0, self.log_message, f"Ошибок: {error_count}")
            self.root.after(0, messagebox.showinfo, "Готово", final_message)

        finally:
            self.root.after(0, self.finish_copying)

    def finish_copying(self):
        self.is_copying = False
        self.btn_source.config(state=NORMAL)
        self.btn_destination.config(state=NORMAL)


if __name__ == "__main__":
    root = Tk()
    app = ExcelFolderCopierApp(root)
    root.mainloop()

# pyinstaller --onefile --windowed --name="Copy all excels" copy_all_excel_folders.py 