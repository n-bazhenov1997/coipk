import xlwings as xw
from docx import Document
import tkinter as tk
from tkinter import messagebox
from typing import Tuple, Dict, List
from datetime import date
from scripts import check_selection


# === 1. Для создания приказа собираем всю необходимую информацию из реестра ===
def collect_info() -> Tuple[Dict[str: str | date], List[str]]:

    # Собираем информацию из выделенного диапазона строк
    book = xw.books.active
    sel = book.selection
    
    # Проверяем: если не выделена хоть одна строка полностью или смежный диапазон строк, то выходим
    if check_selection:
        return

    # Проверяем, все ли номера приказов одинаковы. Если нет, то выводим сообщение об ошибке
    for i in range(sel.rows.count):

        if i == 0:
            ord_num = sel[0, 2].value

        elif sel[i, 2].value != ord_num:
            
            root = tk.Tk()
            root.withdraw()

            messagebox.showerror('Ошибка', 'Один из номеров приказа в выделенном диапазоне не соответствует другим номерам')
            return
        
    # Если номера приказов у всех обучающихся одинаковы, то продолжаем выполнение программы
    
    # Запоминаем базовую информацию 
    main_info = {}
    main_info['ord_date'] = sel[0, 0].value
    main_info['ord_num'] = ord_num
    main_info['prog'] = sel[0, 4].value
    main_info['ord_type'] = sel[0, 9].value
    main_info['prot_num'] = sel[0, 10].value

    # Запоминаем имена студентов
    studs_names = []
    for i in range(sel.rows.count):
        studs_names.append(sel[i, 11].value)

    # Возвращаем информацию о приказе и список имен студентов
    return main_info, studs_names  # Подумай о возвращении еще и ПУТИ К ШАБЛОНУ

def def_doctype():
    pass

        

        
        
    