# -*- coding: utf-8 -*-

from resources import *

import tkinter as tk
from tkinter import ttk

class MainWindow(Window):
    def __init__(self):
        super().__init__() 

    def check_add_prop(self):
        func = PropertyManager()
        func.check_add_properties()

    def get_pdf_window(self):
        window = PDFWindow()
        window.getWindow()

    def initialize_frames(self, root):
        '''
        Метод создания FRAME для окна Tkinter.
        Здесь в словаре заложены все необходимые параметры для создания
        '''
        frames = {
            'main': ttk.Frame(root, borderwidth=10),
            'assembly': ttk.LabelFrame(root, borderwidth=5, relief='solid', 
                                       text='Модуль для работы со сборкой'),
            'drawing': ttk.LabelFrame(root, borderwidth=5, relief='solid', 
                                      text='Модуль для работы с чертежом'),
            'extra': ttk.LabelFrame(root, borderwidth=5, relief='solid', 
                                  text='Дополнительно'),
        }
        
        frames['main'].grid(row=0, column=0, columnspan=2, pady=5, padx=5)
        frames['assembly'].grid(row=1, column=0, pady=5, padx=5, sticky='nsew')
        frames['drawing'].grid(row=2, column=0, pady=5, padx=5, sticky='nsew')
        frames['extra'].grid(row=3, column=0, pady=5, padx=5, sticky='nsew')

        return frames
    
    def skip(self):
        pass
    
    def initialize_buttons(self, frames):
        '''
        Метод создания кнопок. В словаре заложены все необходимые параметры.
        Важно: в словаре кнопки идут в том порядке, в котором они отображаются,
        поэтому при добавлении новой кнопки нужно разместить запись в нужном месте.
        '''

        button_config = [
            {'text': 'Адаптировать сборку насквозь', 'frame': 'assembly', 
             'command': self.skip, 'state': 'disabled'},
            {'text': 'Адаптировать активную сборку', 'frame': 'assembly', 
             'command': AdaptAssy, 'state': 'normal'},
            {'text': 'Адаптировать активную деталь', 'frame': 'assembly', 
             'command': AdaltDetail, 'state': 'normal'},
            {'text': 'Добавить свойства ISO-RGSH', 'frame': 'assembly', 
             'command': self.check_add_prop, 'state': 'normal'},
            {'text': 'Добавить BOM', 'frame': 'drawing', 
             'command': BOMMaker, 'state': 'normal'},
            {'text': 'Создать МТО', 'frame': 'extra', 
             'command': MTOMaker, 'state': 'normal'},
            {'text': 'Установить позиции в активной сборке', 'frame': 'drawing', 
             'command': SetPositions, 'state': 'normal'},
            {'text': 'Добавить тех. требования', 'frame': 'drawing', 
             'command': TechDemandWindow, 'state': 'normal'},
            {'text': 'Добавить таблицу гибов', 'frame': 'drawing', 
             'command': BTWindow, 'state': 'normal'},
            {'text': 'Перевести .CDW чертежи ENG-RUS', 'frame': 'drawing',
              'command': TranslateCDW, 'state': 'normal'},
            {'text': 'Редактор словаря', 'frame': 'extra', 
             'command': DictionaryWindow, 'state': 'normal'},
            {'text': 'Создать PDF', 'frame': 'extra', 
             'command': self.get_pdf_window, 'state': 'normal'}
        ]

        buttons = []
        for i, config in enumerate(button_config):
            frame = frames[config['frame']]
            row = i 
            col = 0
            buttons.append(self.create_button(ttk, frame, 
                                              config['text'], 
                                              config['command'], 
                                              40, 
                                              config['state'], 
                                              row, col))

        return buttons
    
    def get_main_window(self):
        root = tk.Tk()
        root.title(self.window_name)
        root.resizable(False, False)
        root.iconbitmap(self.pic_path)
        root.attributes("-topmost", True)
        
        frames = self.initialize_frames(root)
        buttons = self.initialize_buttons(frames)
        
        root.update_idletasks()
        w, h = self.get_center_window(root)
        root.geometry('+{}+{}'.format(w,h))
        root.mainloop()
