from tkinter import messagebox

import xlwings as xw

def check_selection(is_one_area: bool = True) -> bool:

    try:
        book = xw.books.active
        sheet = book.sheets.active
        sel = book.selection

        if sel.columns.count != sheet.cells.columns.count:
            messagebox.showerror('Ошибка', 'Выделите хотя бы одну строку целиком')
            return False

        elif is_one_area and sel.api.Areas.Count > 1:
            messagebox.showerror('Ошибка', 'Выделите смежный диапазон строк')
            return False

        else:
            return True
    except Exception as e:
        messagebox.showerror('Ошибка', f'Не удалось проверить выделение: {e}')
        return False
