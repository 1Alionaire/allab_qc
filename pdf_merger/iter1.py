# -*- coding: utf-8 -*-
"""
Объединение PDF — настольное приложение (tkinter + PyMuPDF).

Окно разделено на две части:
    слева  — лента миниатюр всех страниц (сюда перетаскиваются PDF с рабочего стола);
    справа — просмотр выбранной страницы.

Горячие клавиши:
    ↑ / ↓            — предыдущая / следующая страница
    Ctrl+↑ / Ctrl+↓  — переместить страницу выше / ниже
    Delete           — удалить выбранную страницу
    Ctrl+Z           — вернуть последнюю удалённую страницу
    Ctrl+O           — добавить файлы через диалог
    Ctrl+S           — экспорт единого PDF

Целевая среда: Windows 7, Python 3.8, PyMuPDF 1.22.5, PyInstaller 4.10.
"""

import logging
import os
import sys
import tempfile
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Новые версии PyMuPDF импортируются как "pymupdf", старые (1.22.5 для Win7) — как "fitz".
# Так код работает и в сборке под Windows 7, и при отладке на свежем Python.
try:
    import pymupdf as fitz
except ImportError:
    import fitz


# ============================================================================
#  НАСТРОЙКИ
# ============================================================================
APP_TITLE = "Объединение PDF"
WINDOW_SIZE = "1200x800"

THUMB_WIDTH = 150           # ширина миниатюры в ленте, px
THUMB_MAX_HEIGHT = 210      # макс. высота миниатюры (для очень вытянутых страниц), px
THUMB_PAD = 14              # отступ между миниатюрами, px
CAPTION_HEIGHT = 34         # высота подписи под миниатюрой (2 строки), px
CAPTION_MAX_CHARS = 24      # длинные имена файлов обрезаются до стольких символов
THUMBS_PER_TICK = 4         # сколько миниатюр рисовать за один шаг (чтобы окно не зависало)

VIEW_MARGIN = 16            # поле вокруг страницы в области просмотра, px
VIEW_MAX_SCALE = 4.0        # максимальное увеличение при просмотре
RESIZE_DELAY_MS = 150       # пауза перед перерисовкой при изменении размера окна, мс
UNDO_LIMIT = 50             # сколько удалений можно отменить через Ctrl+Z

DEFAULT_OUTPUT_NAME = "merged.pdf"
LOG_FILE_NAME = "pdf_merger.log"

COLOR_TAPE_BG = "#e9ecef"   # фон ленты
COLOR_VIEW_BG = "#7f8790"   # фон области просмотра
COLOR_SELECTED = "#1e6fd9"  # рамка выбранной страницы
COLOR_PAGE_BORDER = "#9aa0a6"
FONT_CAPTION = ("Segoe UI", 8)
FONT_HINT = ("Segoe UI", 11)

# Виртуальные коды клавиш Windows. Они НЕ зависят от раскладки,
# поэтому Ctrl+Z / Ctrl+O / Ctrl+S работают и при русской раскладке.
VK_O, VK_S, VK_Z = 79, 83, 90


# ============================================================================
#  ЖУРНАЛ ОШИБОК
# ============================================================================
# В собранном .exe __file__ указывает во временную папку, поэтому журнал кладём рядом с .exe
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def setup_logging():
    """Пишет ошибки в pdf_merger.log рядом с программой (или во временную папку, если нет прав)."""
    for folder in (BASE_DIR, tempfile.gettempdir()):
        path = os.path.join(folder, LOG_FILE_NAME)
        try:
            # delay=True — файл создаётся только при первой записи
            handler = logging.FileHandler(path, encoding="utf-8", delay=True)
            # Проверяем, что в папку реально можно писать
            with open(path, "a", encoding="utf-8"):
                pass
        except OSError:
            continue
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        logging.getLogger().addHandler(handler)
        logging.getLogger().setLevel(logging.INFO)
        return path
    return "(журнал недоступен)"


LOG_PATH = setup_logging()


# ============================================================================
#  DRAG-AND-DROP ИЗ ПРОВОДНИКА WINDOWS (чистый ctypes, без доп. библиотек)
# ============================================================================
def enable_file_drop(widget, on_drop):
    """
    Разрешает перетаскивать файлы из Проводника / с рабочего стола на widget.

    Как это работает: Windows присылает окну сообщение WM_DROPFILES.
    tkinter его не обрабатывает, поэтому мы «подменяем» оконную процедуру
    виджета своей функцией: она ловит WM_DROPFILES, а все остальные
    сообщения передаёт старой процедуре Tk.

    on_drop(paths, x, y) — paths: список путей, x/y: точка «бросания» внутри виджета.
    Возвращает True, если перетаскивание включено (только Windows).
    """
    if sys.platform != "win32":
        return False

    try:
        import ctypes
        from ctypes import wintypes

        # Собственные экземпляры DLL, чтобы наши argtypes не мешали другим библиотекам
        user32 = ctypes.WinDLL("user32")
        shell32 = ctypes.WinDLL("shell32")

        WM_DROPFILES = 0x0233
        GWL_WNDPROC = -4

        LRESULT = wintypes.LPARAM  # знаковое целое размером с указатель
        WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT,
                                     wintypes.WPARAM, wintypes.LPARAM)

        # В 64-битной Windows функция называется SetWindowLongPtrW,
        # в 32-битной её нет — там используется SetWindowLongW.
        set_window_long = getattr(user32, "SetWindowLongPtrW", None) or user32.SetWindowLongW
        set_window_long.restype = ctypes.c_void_p
        set_window_long.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]

        user32.CallWindowProcW.restype = LRESULT
        user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT,
                                           wintypes.WPARAM, wintypes.LPARAM]

        shell32.DragAcceptFiles.restype = None
        shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
        shell32.DragQueryFileW.restype = wintypes.UINT
        shell32.DragQueryFileW.argtypes = [ctypes.c_void_p, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
        shell32.DragQueryPoint.restype = wintypes.BOOL
        shell32.DragQueryPoint.argtypes = [ctypes.c_void_p, ctypes.POINTER(wintypes.POINT)]
        shell32.DragFinish.restype = None
        shell32.DragFinish.argtypes = [ctypes.c_void_p]

        hwnd = widget.winfo_id()  # дескриптор окна (HWND) самого виджета
        old_proc = None           # сюда сохраним оригинальную процедуру Tk (заполняется ниже)

        def window_proc(hwnd_, msg, wparam, lparam):
            if msg != WM_DROPFILES:
                # Всё, что не перетаскивание, — обрабатывает Tk как обычно
                return user32.CallWindowProcW(old_proc, hwnd_, msg, wparam, lparam)

            hdrop = wparam
            try:
                count = shell32.DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)  # сколько файлов
                paths = []
                for i in range(count):
                    length = shell32.DragQueryFileW(hdrop, i, None, 0)
                    buf = ctypes.create_unicode_buffer(length + 1)
                    shell32.DragQueryFileW(hdrop, i, buf, length + 1)
                    paths.append(buf.value)
                point = wintypes.POINT()
                shell32.DragQueryPoint(hdrop, ctypes.byref(point))
                # Тяжёлую работу (открытие PDF) делаем не здесь, а чуть позже в обычном цикле Tk
                widget.after_idle(on_drop, paths, point.x, point.y)
            except Exception:
                logging.error("Ошибка при перетаскивании файлов:\n%s", traceback.format_exc())
            finally:
                shell32.DragFinish(hdrop)  # обязательно освобождаем память Windows
            return 0

        new_proc = WNDPROC(window_proc)
        old_proc = set_window_long(hwnd, GWL_WNDPROC, ctypes.cast(new_proc, ctypes.c_void_p))
        shell32.DragAcceptFiles(hwnd, True)

        # ВАЖНО: держим ссылку на функцию, иначе сборщик мусора её удалит и программа упадёт
        widget._drop_window_proc = new_proc
        return True
    except Exception:
        logging.error("Не удалось включить drag-and-drop:\n%s", traceback.format_exc())
        return False


# ============================================================================
#  ПРИЛОЖЕНИЕ
# ============================================================================
class PdfMergerApp:
    def __init__(self, root):
        self.root = root

        # Каждая страница ленты — словарь:
        #   key      — ключ открытого документа в self.docs
        #   name     — имя файла (для подписи)
        #   page_no  — номер страницы в исходном файле (с 0)
        #   scale    — масштаб для миниатюры
        #   tw, th   — размер миниатюры в пикселях
        #   thumb    — tk.PhotoImage миниатюры; None = ещё не готова; False = ошибка рендера
        #   img_item — id картинки на канве ленты (чтобы подставить миниатюру без полной перерисовки)
        self.entries = []
        self.docs = {}            # нормализованный путь -> открытый документ fitz
        self.selected = None      # индекс выбранной страницы или None
        self.undo_stack = []      # удалённые страницы для Ctrl+Z: (индекс, запись)
        self.tape_boxes = []      # вертикальные границы каждой миниатюры на ленте: (y0, y1)
        self.tape_height = 1      # полная высота содержимого ленты
        self.view_image = None    # PhotoImage текущей страницы (ссылку держать обязательно!)

        self.view_mode = tk.StringVar(value="page")  # "page" — целиком, "width" — по ширине
        self.status = tk.StringVar()

        self._thumb_job = None    # id отложенной задачи рендера миниатюр
        self._view_job = None     # id отложенной перерисовки просмотра
        self._tape_job = None     # id отложенной перерисовки ленты

        self._build_ui()
        self._bind_keys()

        # Перехватываем все необработанные ошибки в обработчиках tkinter:
        # в .exe без консоли иначе они просто «тихо» исчезнут
        root.report_callback_exception = self.on_tk_error

        self.drop_enabled = enable_file_drop(self.tape, self.on_files_dropped)
        self.redraw_tape()
        self.show_current()
        self._update_status()

    # ------------------------------------------------------------------ UI --
    def _build_ui(self):
        root = self.root
        root.title(APP_TITLE)
        root.geometry(WINDOW_SIZE)
        root.minsize(700, 450)

        # --- Панель кнопок ---
        bar = ttk.Frame(root, padding=(6, 6, 6, 0))
        bar.pack(side="top", fill="x")

        ttk.Button(bar, text="Добавить PDF…", command=self.add_files_dialog).pack(side="left")
        ttk.Button(bar, text="▲ Выше", command=lambda: self.move_selected(-1)).pack(side="left", padx=(12, 0))
        ttk.Button(bar, text="▼ Ниже", command=lambda: self.move_selected(+1)).pack(side="left", padx=(4, 0))
        ttk.Button(bar, text="Удалить страницу", command=self.delete_selected).pack(side="left", padx=(12, 0))
        ttk.Button(bar, text="Вернуть", command=self.undo_delete).pack(side="left", padx=(4, 0))
        ttk.Button(bar, text="Очистить всё", command=self.clear_all).pack(side="left", padx=(12, 0))

        ttk.Button(bar, text="Экспорт PDF…", command=self.export_pdf).pack(side="right")
        ttk.Radiobutton(bar, text="По ширине", value="width", variable=self.view_mode,
                        command=self.show_current).pack(side="right", padx=(0, 16))
        ttk.Radiobutton(bar, text="Страница целиком", value="page", variable=self.view_mode,
                        command=self.show_current).pack(side="right", padx=(0, 6))

        # --- Строка состояния (внизу) ---
        ttk.Label(root, textvariable=self.status, anchor="w", padding=(8, 2)).pack(side="bottom", fill="x")

        # --- Рабочая область: лента слева, просмотр справа ---
        paned = ttk.Panedwindow(root, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=6, pady=6)

        left = ttk.Frame(paned)
        self.tape = tk.Canvas(left, width=THUMB_WIDTH + 2 * THUMB_PAD + 16, bg=COLOR_TAPE_BG,
                              highlightthickness=0, yscrollincrement=20)
        tape_scroll = ttk.Scrollbar(left, orient="vertical", command=self.tape.yview)
        self.tape.configure(yscrollcommand=tape_scroll.set)
        tape_scroll.pack(side="right", fill="y")
        self.tape.pack(side="left", fill="both", expand=True)
        paned.add(left, weight=0)

        right = ttk.Frame(paned)
        self.viewer = tk.Canvas(right, bg=COLOR_VIEW_BG, highlightthickness=0,
                                yscrollincrement=20, xscrollincrement=20)
        view_vscroll = ttk.Scrollbar(right, orient="vertical", command=self.viewer.yview)
        view_hscroll = ttk.Scrollbar(right, orient="horizontal", command=self.viewer.xview)
        self.viewer.configure(yscrollcommand=view_vscroll.set, xscrollcommand=view_hscroll.set)
        self.viewer.grid(row=0, column=0, sticky="nsew")
        view_vscroll.grid(row=0, column=1, sticky="ns")
        view_hscroll.grid(row=1, column=0, sticky="ew")
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)
        paned.add(right, weight=1)

        # --- События мыши и изменения размеров ---
        self.tape.bind("<Button-1>", self.on_tape_click)
        self.tape.bind("<Configure>", lambda e: self._schedule_tape_redraw())
        self.viewer.bind("<Configure>", lambda e: self._schedule_view_redraw())
        root.bind_all("<MouseWheel>", self.on_mouse_wheel)

    def _bind_keys(self):
        r = self.root
        r.bind("<Up>", lambda e: self.select_relative(-1))
        r.bind("<Down>", lambda e: self.select_relative(+1))
        r.bind("<Control-Up>", lambda e: self.move_selected(-1))
        r.bind("<Control-Down>", lambda e: self.move_selected(+1))
        r.bind("<Delete>", lambda e: self.delete_selected())
        # Буквенные сочетания ловим по коду клавиши — см. VK_* в настройках
        r.bind("<Control-KeyPress>", self.on_ctrl_key)

    def on_ctrl_key(self, event):
        if event.keycode == VK_Z:
            self.undo_delete()
        elif event.keycode == VK_O:
            self.add_files_dialog()
        elif event.keycode == VK_S:
            self.export_pdf()

    def on_mouse_wheel(self, event):
        # Прокручиваем ту область, над которой находится курсор
        widget = self.root.winfo_containing(event.x_root, event.y_root)
        if widget is self.tape or widget is self.viewer:
            direction = -1 if event.delta > 0 else 1
            widget.yview_scroll(direction * 3, "units")

    def on_tk_error(self, exc_type, exc_value, exc_tb):
        text = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
        logging.error("Необработанная ошибка:\n%s", text)
        messagebox.showerror(APP_TITLE, "Произошла ошибка:\n{}\n\nПодробности в журнале:\n{}".format(
            exc_value, LOG_PATH))

    # ------------------------------------------------- Добавление файлов --
    def add_files_dialog(self):
        paths = filedialog.askopenfilenames(
            title="Выберите PDF-файлы",
            filetypes=[("PDF-файлы", "*.pdf"), ("Все файлы", "*.*")])
        if paths:
            self.add_files(list(paths), insert_at=len(self.entries))

    def on_files_dropped(self, paths, x, y):
        # Вставляем страницы туда, куда «бросили» файлы: между миниатюрами
        insert_at = self._index_for_drop(self.tape.canvasy(y))
        self.add_files(paths, insert_at)

    def _index_for_drop(self, canvas_y):
        for i, (y0, y1) in enumerate(self.tape_boxes):
            if canvas_y < (y0 + y1) / 2:  # выше середины миниатюры — вставляем перед ней
                return i
        return len(self.entries)

    def add_files(self, paths, insert_at):
        pdf_paths, skipped = self._collect_pdf_paths(paths)
        new_entries = []
        errors = []

        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            for path in pdf_paths:
                try:
                    key, doc = self._open_document(path)
                    for page_no in range(doc.page_count):
                        new_entries.append(self._make_entry(key, path, doc, page_no))
                except Exception as exc:
                    logging.error("Не удалось открыть %s:\n%s", path, traceback.format_exc())
                    errors.append("{} — {}".format(os.path.basename(path), exc))
        finally:
            self.root.config(cursor="")

        if new_entries:
            insert_at = max(0, min(insert_at, len(self.entries)))
            self.entries[insert_at:insert_at] = new_entries  # вставка списка в середину
            self.select(insert_at)                            # выделяем первую добавленную
            self._start_thumb_rendering()

        problems = []
        if errors:
            problems.append("Не удалось открыть:\n" + "\n".join(errors))
        if skipped:
            problems.append("Пропущены (не PDF):\n" + "\n".join(skipped))
        if problems:
            messagebox.showwarning(APP_TITLE, "\n\n".join(problems))

    @staticmethod
    def _collect_pdf_paths(paths):
        """Из перетащенных путей выбирает PDF. Если перетащили папку — берём все PDF из неё."""
        pdf_paths, skipped = [], []
        for path in paths:
            if os.path.isdir(path):
                for name in sorted(os.listdir(path), key=str.lower):
                    if name.lower().endswith(".pdf"):
                        pdf_paths.append(os.path.join(path, name))
            elif path.lower().endswith(".pdf"):
                pdf_paths.append(path)
            else:
                skipped.append(os.path.basename(path))
        return pdf_paths, skipped

    def _open_document(self, path):
        """
        Открывает PDF (или возвращает уже открытый).
        Файл читается в память целиком — так он не блокируется на диске,
        и его можно переименовать, удалить или даже перезаписать при экспорте.
        """
        key = os.path.normcase(os.path.abspath(path))
        if key in self.docs:
            return key, self.docs[key]

        with open(path, "rb") as f:
            data = f.read()
        doc = fitz.open(stream=data, filetype="pdf")
        if doc.needs_pass:
            doc.close()
            raise ValueError("файл защищён паролем")
        if doc.page_count == 0:
            doc.close()
            raise ValueError("в файле нет страниц")

        self.docs[key] = doc
        return key, doc

    @staticmethod
    def _make_entry(key, path, doc, page_no):
        rect = doc.load_page(page_no).rect  # размер страницы с учётом её поворота
        page_w = max(rect.width, 1)
        page_h = max(rect.height, 1)
        scale = min(THUMB_WIDTH / page_w, THUMB_MAX_HEIGHT / page_h)
        return {
            "key": key,
            "name": os.path.basename(path),
            "page_no": page_no,
            "scale": scale,
            "tw": max(int(page_w * scale), 1),
            "th": max(int(page_h * scale), 1),
            "thumb": None,
            "img_item": None,
        }

    # ------------------------------------------------------- Рендеринг --
    def _render(self, entry, scale):
        """Рисует страницу в картинку tkinter нужного масштаба."""
        page = self.docs[entry["key"]].load_page(entry["page_no"])
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        # Формат PPM tkinter понимает сам — Pillow не нужен
        return tk.PhotoImage(data=pix.tobytes("ppm"))

    def _start_thumb_rendering(self):
        if self._thumb_job is None:
            self._thumb_job = self.root.after(1, self._render_thumbs_step)

    def _render_thumbs_step(self):
        """Рисует несколько миниатюр за раз и уступает очередь окну — так интерфейс не зависает."""
        self._thumb_job = None
        done = 0
        for entry in self.entries:
            if entry["thumb"] is not None:
                continue
            try:
                entry["thumb"] = self._render(entry, entry["scale"])
            except Exception:
                logging.error("Ошибка рендера миниатюры %s стр.%s:\n%s",
                              entry["name"], entry["page_no"] + 1, traceback.format_exc())
                entry["thumb"] = False  # больше не пытаемся
            # Подставляем картинку в уже нарисованную рамку — без полной перерисовки ленты
            if entry["img_item"] is not None and entry["thumb"]:
                self.tape.itemconfigure(entry["img_item"], image=entry["thumb"])
            done += 1
            if done >= THUMBS_PER_TICK:
                break

        if any(e["thumb"] is None for e in self.entries):
            self._thumb_job = self.root.after(1, self._render_thumbs_step)
        self._update_status()

    # ------------------------------------------------------------ Лента --
    def _schedule_tape_redraw(self):
        if self._tape_job is not None:
            self.root.after_cancel(self._tape_job)
        self._tape_job = self.root.after(RESIZE_DELAY_MS, self._delayed_tape_redraw)

    def _delayed_tape_redraw(self):
        self._tape_job = None
        self.redraw_tape()

    def redraw_tape(self):
        """Полностью перерисовывает ленту миниатюр."""
        canvas = self.tape
        canvas.delete("all")
        self.tape_boxes = []

        canvas_w = canvas.winfo_width()
        if canvas_w <= 1:  # окно ещё не показано — берём запрошенную ширину
            canvas_w = int(canvas["width"])

        if not self.entries:
            canvas.create_text(canvas_w / 2, 60, text="Перетащите сюда\nPDF-файлы",
                               justify="center", font=FONT_HINT, fill="#6c757d")
            self.tape_height = 1
            canvas.configure(scrollregion=(0, 0, canvas_w, 1))
            return

        y = THUMB_PAD
        for i, entry in enumerate(self.entries):
            tw, th = entry["tw"], entry["th"]
            x0 = (canvas_w - tw) / 2

            if i == self.selected:
                canvas.create_rectangle(x0 - 5, y - 5, x0 + tw + 5, y + th + 5,
                                        fill=COLOR_SELECTED, outline=COLOR_SELECTED)
            # Белая «страница» — рамка и заглушка, пока миниатюра не готова
            canvas.create_rectangle(x0, y, x0 + tw, y + th, fill="white", outline=COLOR_PAGE_BORDER)
            if entry["thumb"] is False:
                canvas.create_text(x0 + tw / 2, y + th / 2, text="ошибка", fill="red")
            entry["img_item"] = canvas.create_image(x0, y, anchor="nw", image=entry["thumb"] or "")

            name = entry["name"]
            if len(name) > CAPTION_MAX_CHARS:
                name = name[:CAPTION_MAX_CHARS - 1] + "…"
            caption = "{}\n{} · стр. {}".format(i + 1, name, entry["page_no"] + 1)
            canvas.create_text(canvas_w / 2, y + th + 4, text=caption, anchor="n", justify="center",
                               font=FONT_CAPTION,
                               fill=COLOR_SELECTED if i == self.selected else "#343a40")

            bottom = y + th + CAPTION_HEIGHT
            self.tape_boxes.append((y - THUMB_PAD / 2, bottom + THUMB_PAD / 2))
            y = bottom + THUMB_PAD

        self.tape_height = y
        canvas.configure(scrollregion=(0, 0, canvas_w, y))

    def _scroll_tape_to(self, index):
        """Прокручивает ленту так, чтобы страница index была видна."""
        if not self.tape_boxes or self.tape_height <= 1:
            return
        y0, y1 = self.tape_boxes[index]
        view_top = self.tape.canvasy(0)
        view_h = self.tape.winfo_height()
        if view_h <= 1:  # окно ещё не показано на экране
            return
        if y0 < view_top:
            self.tape.yview_moveto(y0 / self.tape_height)
        elif y1 > view_top + view_h:
            self.tape.yview_moveto((y1 - view_h) / self.tape_height)

    def on_tape_click(self, event):
        self.tape.focus_set()  # убираем фокус с кнопок, чтобы стрелки работали с лентой
        canvas_y = self.tape.canvasy(event.y)
        for i, (y0, y1) in enumerate(self.tape_boxes):
            if y0 <= canvas_y <= y1:
                self.select(i)
                return

    # -------------------------------------------------------- Просмотр --
    def _schedule_view_redraw(self):
        if self._view_job is not None:
            self.root.after_cancel(self._view_job)
        self._view_job = self.root.after(RESIZE_DELAY_MS, self._delayed_view_redraw)

    def _delayed_view_redraw(self):
        self._view_job = None
        self.show_current(reset_scroll=False)

    def show_current(self, reset_scroll=True):
        """Показывает выбранную страницу в правой части с нужным масштабом."""
        canvas = self.viewer
        canvas.delete("all")
        self.view_image = None
        canvas_w = max(canvas.winfo_width(), 100)
        canvas_h = max(canvas.winfo_height(), 100)

        if self.selected is None:
            hint = "Перетащите PDF-файлы на ленту слева\nили нажмите «Добавить PDF…»"
            canvas.create_text(canvas_w / 2, canvas_h / 2, text=hint, justify="center",
                               font=FONT_HINT, fill="white")
            canvas.configure(scrollregion=(0, 0, canvas_w, canvas_h))
            return

        entry = self.entries[self.selected]
        try:
            rect = self.docs[entry["key"]].load_page(entry["page_no"]).rect
            avail_w = canvas_w - 2 * VIEW_MARGIN
            avail_h = canvas_h - 2 * VIEW_MARGIN
            if self.view_mode.get() == "width":
                scale = avail_w / rect.width
            else:
                scale = min(avail_w / rect.width, avail_h / rect.height)
            scale = max(0.05, min(scale, VIEW_MAX_SCALE))
            self.view_image = self._render(entry, scale)
        except Exception:
            logging.error("Ошибка просмотра %s стр.%s:\n%s",
                          entry["name"], entry["page_no"] + 1, traceback.format_exc())
            canvas.create_text(canvas_w / 2, canvas_h / 2, text="Не удалось отобразить страницу",
                               font=FONT_HINT, fill="white")
            return

        img_w, img_h = self.view_image.width(), self.view_image.height()
        total_w = max(img_w + 2 * VIEW_MARGIN, canvas_w)
        total_h = max(img_h + 2 * VIEW_MARGIN, canvas_h)
        canvas.create_image(total_w / 2, VIEW_MARGIN, anchor="n", image=self.view_image)
        canvas.configure(scrollregion=(0, 0, total_w, total_h))
        if reset_scroll:
            canvas.xview_moveto(0)
            canvas.yview_moveto(0)

    # ----------------------------------------------- Выбор и изменения --
    def select(self, index):
        if not self.entries:
            self.selected = None
        else:
            self.selected = max(0, min(index, len(self.entries) - 1))
        self.redraw_tape()
        if self.selected is not None:
            self._scroll_tape_to(self.selected)
        self.show_current()
        self._update_status()

    def select_relative(self, step):
        if self.selected is not None:
            self.select(self.selected + step)

    def move_selected(self, step):
        """Меняет выбранную страницу местами с соседней (step = -1 выше, +1 ниже)."""
        i = self.selected
        if i is None:
            return
        j = i + step
        if not 0 <= j < len(self.entries):
            return
        self.entries[i], self.entries[j] = self.entries[j], self.entries[i]
        self.selected = j
        self.redraw_tape()
        self._scroll_tape_to(j)
        self._update_status()

    def delete_selected(self):
        if self.selected is None:
            return
        entry = self.entries.pop(self.selected)
        self.undo_stack.append((self.selected, entry))
        if len(self.undo_stack) > UNDO_LIMIT:
            self.undo_stack.pop(0)
        self.select(self.selected)  # select() сам поправит индекс, если удалили последнюю

    def undo_delete(self):
        if not self.undo_stack:
            return
        index, entry = self.undo_stack.pop()
        index = min(index, len(self.entries))
        self.entries.insert(index, entry)
        self.select(index)

    def clear_all(self):
        if not self.entries:
            return
        if not messagebox.askyesno(APP_TITLE, "Убрать все страницы из ленты?"):
            return
        if self._thumb_job is not None:
            self.root.after_cancel(self._thumb_job)
            self._thumb_job = None
        self.entries = []
        self.undo_stack = []  # удалённые страницы ссылаются на закрываемые документы
        for doc in self.docs.values():
            doc.close()
        self.docs = {}
        self.select(0)

    # ----------------------------------------------------------- Экспорт --
    def export_pdf(self):
        if not self.entries:
            messagebox.showinfo(APP_TITLE, "Лента пуста — нечего сохранять.")
            return
        out_path = filedialog.asksaveasfilename(
            title="Сохранить единый PDF",
            defaultextension=".pdf",
            initialfile=DEFAULT_OUTPUT_NAME,
            filetypes=[("PDF-файлы", "*.pdf")])
        if not out_path:
            return

        self.status.set("Сохранение…")
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        out = fitz.open()  # новый пустой документ
        try:
            # Подряд идущие страницы одного файла копируем одним блоком — так быстрее
            i, n = 0, len(self.entries)
            while i < n:
                first = self.entries[i]
                j = i
                while (j + 1 < n
                       and self.entries[j + 1]["key"] == first["key"]
                       and self.entries[j + 1]["page_no"] == self.entries[j]["page_no"] + 1):
                    j += 1
                out.insert_pdf(self.docs[first["key"]],
                               from_page=first["page_no"], to_page=self.entries[j]["page_no"])
                i = j + 1
            # garbage=3 убирает дубли общих ресурсов (шрифты и т.п.), deflate=True сжимает потоки
            out.save(out_path, garbage=3, deflate=True)
        except Exception as exc:
            logging.error("Ошибка экспорта в %s:\n%s", out_path, traceback.format_exc())
            messagebox.showerror(APP_TITLE, "Не удалось сохранить файл:\n{}\n\n"
                                            "Если он открыт в другой программе — закройте её.".format(exc))
            return
        finally:
            out.close()
            self.root.config(cursor="")
            self._update_status()

        if messagebox.askyesno(APP_TITLE, "Сохранено страниц: {}\n{}\n\nОткрыть файл?".format(
                len(self.entries), out_path)):
            if hasattr(os, "startfile"):  # есть только в Windows
                os.startfile(out_path)

    # ---------------------------------------------------- Строка статуса --
    def _update_status(self):
        parts = ["Страниц в ленте: {}".format(len(self.entries))]
        if self.selected is not None:
            e = self.entries[self.selected]
            total = self.docs[e["key"]].page_count
            parts.append("Выбрана: {} — {}, стр. {} из {}".format(
                self.selected + 1, e["name"], e["page_no"] + 1, total))
        if any(e["thumb"] is None for e in self.entries):
            parts.append("готовятся миниатюры…")
        if not self.drop_enabled:
            parts.append("перетаскивание недоступно — используйте «Добавить PDF…»")
        self.status.set("   |   ".join(parts))


def main():
    root = tk.Tk()
    app = PdfMergerApp(root)
    # Если PDF перетащили прямо на значок .exe — Windows передаст пути в аргументах
    if len(sys.argv) > 1:
        root.after(200, lambda: app.add_files(sys.argv[1:], insert_at=0))
    root.mainloop()


if __name__ == "__main__":
    main()

#PyMuPDF==1.22.5  
# python -m pip install -r requirements.txt
# python -m PyInstaller --onefile --windowed --name PDF_Merger pdf_merger.py