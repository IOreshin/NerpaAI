# -*- coding: utf-8 -*-

from resources import *

import tkinter as tk
from tkinter import ttk

import sys


class WalrusWindow(Window):
    def __init__(self):
        super().__init__()
        self.root = tk.Tk()
        self.root.title('WalrusAI')
        self.root.resizable(False, False)
        self.root.iconbitmap(self.pic_path)
        self.root.attributes("-topmost", True)

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.initialize_ui()

    def on_closing(self):
        self.root.destroy()
        sys.exit()

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
            'extra': ttk.LabelFrame(root, borderwidth=5, relief='solid', 
                                  text='Функции'),
        }
        
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
            {'text': 'Перевести .CDW чертежи RUS-ENG', 'frame': 'extra',
              'command': RuEnTranslateCDW, 'state': 'normal'},
            {'text': 'Редактор словаря', 'frame': 'extra', 
             'command': DictionaryWindow, 'state': 'normal'},
            {'text': 'Создать PDF', 'frame': 'extra', 
             'command': self.get_pdf_window, 'state': 'normal'},
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
    
    def initialize_ui(self):
        frames = self.initialize_frames(self.root)
        buttons = self.initialize_buttons(frames)
        
    def run(self):
        self.root.update_idletasks()
        w, h = self.get_center_window(self.root)
        self.root.geometry('+{}+{}'.format(w, h))
        self.root.mainloop()

if __name__ == '__main__':
    app = MainWindow()
    app.run()