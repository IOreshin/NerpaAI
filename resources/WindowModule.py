# -*- coding: utf-8 -*-

from .NerpaUtility import get_path
import tkinter as tk

class Window:
    '''
    Базовый класс для всех окон приложения NerpaAI
    '''
    def __init__(self):
        self.window_name = 'NerpaAI'
        self.pic_path = get_path()+'\\pic\\2_old.ico'
        
    def create_button(self, ttk, frame, text, command, width, state, row, column):
        '''
        Функция для создания кнопки
        '''
        button = ttk.Button(frame, text = text, command = command, 
                            width = width, state = state)
        button.grid(row = row, column = column, padx = 5, pady = 5)
        return button
    
    def get_center_window(self, root):
        '''
        Функция автоматического размещения окна по центру экрана
        '''
        s = root.geometry()
        s = s.split('+')
        s = s[0].split('x')
        w_root = int(s[0])
        h_root = int(s[1])

        w_screen = root.winfo_screenwidth()
        h_screen = root.winfo_screenheight()

        w = (w_screen - w_root) // 2
        h = (h_screen - h_root) // 2

        return w,h
    
