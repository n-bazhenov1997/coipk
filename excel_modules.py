import xlwings as xw
import tkinter as tk
from tkinter import simpledialog, messagebox
from pathlib import Path

root = tk.Tk()
root.withdraw()

def find_protocol_paths():
    book = xw.books.active
    sheet = book.sheets.active
    selection = book.selection

    # Если выделена полностью хоть одна строка
    if selection.columns.count == sheet.cells.columns.count:

        # Запоминаем: выделение - правда, рабочая область - выделение
        selection = selection.api

        if selection.Areas.Count > 1:
            messagebox.showerror('Ошибка', 'Выделите только ОДНУ ОБЛАСТЬ для корректной работы или не выделяйте НИЧЕГО (тогда будет проанализирован ВЕСЬ рабочий диапазон)')
            return
        else:
            is_selection = True
            working_area = book.selection

        # Запрашиваем у пользователя путь к папке с протоколами
        path = simpledialog.askstring(
            'Ввод данных', 
            'Введите путь к папке с выделенными протоколами'
            )

        # Если пользователь ничего не передал - выходим
        if path is None:
            return
        
        # Если путь введен, то преобразуем его в объект и проверяем
        else:
            path = Path(path)
            if not path.is_dir():
                messagebox.showerror('Ошибка', 'Введен неверный путь')
                return

    # Если не выделена ни одна строка
    else:

        is_selection = False
        working_area = sheet.used_range
