# -*- coding: utf-8 -*-

from WindowModule import Window
from PositionModule import SetPositions
from PropertyMngModule import PropertyManager
from TechDemandsModule import TechDemandWindow
from AdaptModule import AdaltDetail, AdaptAssy
from BendModule import BTWindow
from ReportModule import MTOMaker, BOMMaker
from TranslateModule import TranslateCDW
from DictionaryModule import DictionaryWindow
from PDFModule import PDFWindow

import tkinter as tk
from tkinter import ttk
from tkinter import font

class MainWindow(Window):
    def __init__(self):
        super().__init__()

    def adapt_assy(self):
        func = AdaptAssy()
        func.adapt_current_assy()

    def adapt_detail(self):
        func = AdaltDetail()
        func.adapt_current_detail()

    def bend_table(self):
        wind = BTWindow()
        wind.get_bt_window()

    def get_BOM(self):
        func = BOMMaker()
        func.place_bom_drw()

    def get_MTO(self):
        func = MTOMaker()
        func.get_report()

    def tech_demand_window(self):
        wind = TechDemandWindow()
        wind.get_tt_window()

    def check_add_prop(self):
        func = PropertyManager()
        func.check_add_properties()

    def set_positions(self):
        func = SetPositions()
        func.set_positions()
        
    def translate_cdw(self):
        func = TranslateCDW()
        func.get_cdw_docs()

    def dict_editor(self):
        func = DictionaryWindow()
        func.get_dictionary_window()

    def get_pdf_window(self):
        window = PDFWindow()
        window.getWindow()

    def get_main_window(self):
        root = tk.Tk()
        root.title(self.window_name)
        root.resizable(False, False)
        root.iconbitmap(self.pic_path)
        root.attributes("-topmost", True)
        font1 = font.Font(family= "yu gothic ui", size = 10, weight = "bold")

        frame1 = ttk.Frame(root, borderwidth = 10)
        frame2 = ttk.LabelFrame(root, borderwidth = 5, 
                                relief = 'solid', text = 'Модуль для работы со сборкой')
        frame3 = ttk.LabelFrame(root, borderwidth = 5, 
                                relief = 'solid', text = 'Модуль для работы с чертежом')
        frame4 = ttk.LabelFrame(root, borderwidth = 5, 
                                relief = 'solid', text = 'Модуль МТО')
        frame5 = ttk.LabelFrame(root, borderwidth = 5, 
                                relief = 'solid', text = 'Модуль PDF')

        frame1.grid(row = 0, column = 0, columnspan = 2, pady = 5, padx = 5)
        frame2.grid(row = 1, column = 0, pady = 5, padx = 5, sticky = 'nsew')
        frame3.grid(row = 2, column = 0, pady = 5, padx = 5, sticky = 'nsew')
        frame4.grid(row = 3, column = 0, pady = 5, padx = 5, sticky = 'nsew')
        frame5.grid(row = 4, column = 0, pady = 5, padx = 5, sticky = 'nsew')

        frames = [frame1, frame2, frame3, frame4, frame5]

        def skip():
            pass


        """В МАССИВЕ ПО ПОРЯДКУ ЗАЛОЖЕНЫ: НАЗВАНИЕ КНОПКИ,
        ПРИНАДЛЕЖНОСТЬ К ФРЕЙМУ, ВЫПОЛНЯЕМАЯ КОМАНДА, СОСТОЯНИЕ"""
        btn_params = [
                ['Адаптировать сборку насквозь',1, skip, ["disabled"]], #31
                ['Адаптировать активную сборку',1, self.adapt_assy, ["normal"]], #2
                ['Адаптировать активную деталь',1, self.adapt_detail, ["normal"]], #3
                ['Добавить свойства ISO-RGSH',1,self.check_add_prop, ["normal"]], #4
                ['Добавить BOM',2, self.get_BOM , ["normal"]], #5
                ['Создать МТО',3, self.get_MTO, ["normal"]], #6
                ['Установить позиции в активной сборке', 2, self.set_positions, ["normal"]], #7
                ['Добавить тех. требования', 2, self.tech_demand_window, ["normal"]], #8
                ['Добавить таблицу гибов', 2, self.bend_table, ["normal"]], #9
                ['Перевести .CDW чертежи ENG-RUS', 2, self.translate_cdw, ["normal"]], #10
                ['Редактор словаря', 2, self.dict_editor, ["normal"]], #11
                ['Создать PDF', 4, self.get_pdf_window, ["normal"]], #12
                ]

        positions = [
                    (0, 0), #1
                    (1, 0), #2
                    (2, 0), #3
                    (3, 0), #4
                    (1, 0), #5
                    (0, 0), #6
                    (0, 0), #7
                    (2, 0), #8
                    (3, 0), #9
                    (4, 0), #10
                    (5, 0), #11
                    (0, 0), #12
                     ] 

        buttons = [self.create_button(ttk, frames[btn_params[i][1]], btn_params[i][0], btn_params[i][2], 40,
                    btn_params[i][3], *positions[i]) for i in range(len(btn_params))]

        root.update_idletasks()
        w,h = self.get_center_window(root)
        root.geometry('+{}+{}'.format(w,h))

        root.mainloop()

