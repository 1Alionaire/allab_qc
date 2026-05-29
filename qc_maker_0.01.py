"""
Обработка одного Excel-файла через win32com.
Сохраняет картинки, макросы и форматирование исходника.

Структура:
  edit_workbook(wb, filepath)  — ТУТ ТВОЯ ЛОГИКА, всё дорабатывай здесь
  process_one_file(filepath)   — обвязка: открыть → вызвать edit_workbook → сохранить
  вспомогательные функции      — sheet_exists, last_row, read_column, write_column
"""

import os
import gc
import pythoncom
import win32com.client as win32


# Константы Excel
XL_AUTOMATION_SECURITY_LOW = 1   # без Protected View, макросы разрешены
XL_UP = -4162                     # для .End(XL_UP) — последняя заполненная строка


# =========================================================
#  ВОТ ТУТ — твоя бизнес-логика. Меняй внутри этой функции.
# =========================================================
def edit_workbook(wb, filepath):
    """
    Получает открытую книгу wb (COM Workbook) и путь к файлу.
    Делает все нужные правки. Save/Close — снаружи.

    Возвращай dict с тем, что хочешь получить наружу — оно дойдёт
    до вызывающего кода в поле 'data'.
    """
    ws = wb.Worksheets('PLM_TEM_Report')

    # --- Пример: записать значение ---
    ws.Range('C3').Value = 'Иванов'

    # --- Пример: прочитать значение ---
    project_code = ws.Range('B5').Value

    # --- Пример: дата ---
    # import datetime as dt
    # ws.Range('C5').Value = dt.date(2026, 5, 27)

    # --- Пример: цикл по диапазону ---
    # for row in range(10, 20):
    #     ws.Cells(row, 3).Value = f'строка {row}'

    # --- Пример: условие ---
    # if ws.Range('D1').Value == 'OV':
    #     ws.Range('E1').Value = 'обработано'

    # --- Пример: работа с другим листом, если есть ---
    # if sheet_exists(wb, 'Summary'):
    #     ws2 = wb.Worksheets('Summary')
    #     ws2.Range('A1').Value = project_code

    return {
        'project_code': project_code,
    }


# =========================================================
#  Обвязка — менять её обычно не нужно
# =========================================================
def process_one_file(filepath, save=True):
    """
    Открывает файл, вызывает edit_workbook, сохраняет, закрывает.

    save=False — для отладки: правки применятся в памяти, но в файл не уйдут.
                 Удобно прогонять, пока edit_workbook сырая.

    Возвращает dict: {'status': 'ok'/'error', 'filepath': ..., 'data': ..., 'error': ...}
    """
    filepath = os.path.abspath(filepath)
    if not os.path.isfile(filepath):
        raise FileNotFoundError(filepath)

    pythoncom.CoInitialize()
    excel = None
    wb = None
    result = {'status': 'error', 'filepath': filepath, 'data': None, 'error': None}
    try:
        excel = _new_excel()
        wb = excel.Workbooks.Open(filepath, 0, False)  # UpdateLinks=0, ReadOnly=False
        if wb is None:
            raise RuntimeError('Workbooks.Open вернул None')

        # === Твоя логика ===
        result['data'] = edit_workbook(wb, filepath)
        # ===================

        if save:
            wb.Save()

        result['status'] = 'ok'

    except Exception as ex:
        result['error'] = str(ex)
        raise
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        gc.collect()
        pythoncom.CoUninitialize()

    return result


def _new_excel():
    """Свежий экземпляр Excel с настройками для тихой работы."""
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    excel.AskToUpdateLinks = False
    excel.AutomationSecurity = XL_AUTOMATION_SECURITY_LOW
    return excel


# =========================================================
#  Хелперы — используй внутри edit_workbook
# =========================================================
def sheet_exists(wb, sheet_name):
    """True, если лист с таким именем есть в книге."""
    try:
        wb.Worksheets(sheet_name)
        return True
    except Exception:
        return False


def get_or_create_sheet(wb, sheet_name):
    """Получить лист, создать если нет."""
    if sheet_exists(wb, sheet_name):
        return wb.Worksheets(sheet_name)
    new_sheet = wb.Worksheets.Add()
    new_sheet.Name = sheet_name
    return new_sheet


def last_row(ws, column=1):
    """Номер последней заполненной строки в указанной колонке (1-based)."""
    return ws.Cells(ws.Rows.Count, column).End(XL_UP).Row


def read_column(ws, column, start_row=1, end_row=None):
    """
    Прочитать колонку как python-список.
    column — номер колонки (A=1, B=2, ...).
    """
    if end_row is None:
        end_row = last_row(ws, column)
    if end_row < start_row:
        return []
    rng = ws.Range(ws.Cells(start_row, column), ws.Cells(end_row, column))
    raw = rng.Value
    if start_row == end_row:
        return [raw]
    return [row[0] for row in raw]


def write_column(ws, column, start_row, values):
    """Записать список значений в колонку, начиная со start_row."""
    if not values:
        return
    end_row = start_row + len(values) - 1
    rng = ws.Range(ws.Cells(start_row, column), ws.Cells(end_row, column))
    rng.Value = tuple((v,) for v in values)  # Excel ждёт tuple of tuples


def read_row(ws, row, start_col=1, end_col=None):
    """Прочитать строку как python-список."""
    if end_col is None:
        end_col = ws.Cells(row, ws.Columns.Count).End(-4159).Column  # xlToLeft = -4159
    if end_col < start_col:
        return []
    rng = ws.Range(ws.Cells(row, start_col), ws.Cells(row, end_col))
    raw = rng.Value
    if start_col == end_col:
        return [raw]
    return list(raw[0])


# =========================================================
#  Точка входа
# =========================================================
if __name__ == '__main__':
    result = process_one_file(
        r'C:\Python\allab\allab_qc\all_plm\for alibek - Copy\260102-1_PLM_NOB_TEM.xlsm',
        save=True,  # поставь False, пока отлаживаешь edit_workbook
    )
    print(result)