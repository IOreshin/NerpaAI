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

    def get_main_window(self):
        root = tk.Tk()
        root.title(self.window_name)
        root.resizable(False, False)
        root.iconbitmap(self.pic_path)
        root.attributes("-topmost", True)
        font1 = font.Font(family= "yu gothic ui", size = 10, weight = "bold")

        frame1 = ttk.Frame(root, borderwidth = 10)
        frame2 = ttk.LabelFrame(root, borderwidth = 5, relief = 'solid', text = 'ADAPT MODULE')
        frame3 = ttk.LabelFrame(root, borderwidth = 5, relief = 'solid', text = 'DRAWING MODULE')
        frame4 = ttk.LabelFrame(root, borderwidth = 5, relief = 'solid', text = 'MTO MODULE')

        frame1.grid(row = 0, column = 0, columnspan = 2, pady = 5, padx = 5)
        frame2.grid(row = 1, column = 0, pady = 5, padx = 5, sticky = 'nsew')
        frame3.grid(row = 2, column = 0, pady = 5, padx = 5, sticky = 'nsew')
        frame4.grid(row = 3, column = 0, pady = 5, padx = 5, sticky = 'nsew')


        frames = [frame1, frame2, frame3, frame4]

        def skip():
            pass


        """В МАССИВЕ ПО ПОРЯДКУ ЗАЛОЖЕНЫ: НАЗВАНИЕ КНОПКИ,
        ПРИНАДЛЕЖНОСТЬ К ФРЕЙМУ, ВЫПОЛНЯЕМАЯ КОМАНДА, СОСТОЯНИЕ"""
        btn_params = [
                ['FULL ASSY',1, skip, ["disabled"]], #31
                ['CURRENT ASSY',1, self.adapt_assy, ["normal"]], #2
                ['CURRENT DETAIL',1, self.adapt_detail, ["normal"]], #3
                ['ADD PROPERTIES',1, self.check_add_prop, ["normal"]], #4
                ['ADD BOM',2, self.get_BOM , ["normal"]], #5
                ['MAKE MTO',3, self.get_MTO, ["normal"]], #6
                ['SET POSITIONS', 2, self.set_positions, ["normal"]], #7
                ['ADD NOTES', 2, self.tech_demand_window, ["normal"]], #8
                ['ADD BEND TABLE', 2, self.bend_table, ["normal"]], #9
                ['TRANSLATE CDW', 2, skip, ["normal"]], #10
                ['DICTIONARY EDITOR', 2, skip, ["normal"]], #11
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
                    (3, 0),  #9
                    (4, 0),  #9
                    (5, 0),  #9
                    ] 

        buttons = [self.create_button(ttk, frames[btn_params[i][1]], btn_params[i][0], btn_params[i][2], 40,
                    btn_params[i][3], *positions[i]) for i in range(len(btn_params))]

        root.update_idletasks()
        w,h = self.get_center_window(root)
        root.geometry('+{}+{}'.format(w,h))

        root.mainloop()

window = MainWindow()
window.get_main_window()

