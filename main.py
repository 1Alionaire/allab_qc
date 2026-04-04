import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading




class ExcelMainApp:

    def __init__(self, root):
        self.root = root
        self.root.title('Find QC')
        self.root.geometry("1200x800")
        self.output_folder = ""
        self.create_widgets()

    def create_widgets(self):

        self.btn_select = tk.Button(
            self.root, 
            text="Choose Folder", 
            command=self.select_folder,
            width=20,
            height=2
        )
        self.btn_select.pack(pady=10)
        
        # Метка с выбранной папкой
        self.lbl_folder = tk.Label(self.root, text="Папка не выбрана", wraplength=550)
        self.lbl_folder.pack(pady=5)
        
        # Кнопка запуска
        self.btn_run = tk.Button(
            self.root, 
            text="Запустить", 
            command=self.run_processing,
            width=20,
            height=2,
            state=tk.DISABLED
        )
        self.btn_run.pack(pady=10)
        
        # Прогресс-бар
        self.progress = ttk.Progressbar(self.root, length=550, mode='determinate')
        self.progress.pack(pady=10)
        
        # Текущий статус
        self.lbl_status = tk.Label(self.root, text="", font=("Arial", 10))
        self.lbl_status.pack(pady=5)
        
        # Лог
        self.log = tk.Text(self.root, height=10, width=70, state=tk.DISABLED)
        self.log.pack(pady=10)
    
    def select_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с Excel-файлами")
        if folder:
            self.folder_path = folder
            self.lbl_folder.config(text=f"Папка: {folder}")
            self.btn_run.config(state=tk.NORMAL)

    def run_processing(self):
        """Запускает обработку в отдельном потоке"""
        self.btn_run.config(state=tk.DISABLED)
        self.btn_select.config(state=tk.DISABLED)
        self.progress['value'] = 0
        
        # Очищаем лог
        self.log.config(state=tk.NORMAL)
        self.log.delete(1.0, tk.END)
        self.log.config(state=tk.DISABLED)
        
        # Запускаем в отдельном потоке, чтобы GUI не зависал
        thread = threading.Thread(target=self.process_files)
        thread.start()
    
    def process(files)
    
def main():
    root = tk.Tk()
    app = ExcelMainApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()