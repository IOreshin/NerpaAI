# -*- coding: utf-8 -*-

import sqlite3, os, json

import tkinter as tk
from tkinter import ttk
from WindowModule import Window
from NerpaUtility import read_json, get_path
from tkinter.messagebox import showerror, showinfo


class DictionaryWindow(Window):
    def __init__(self):
        super().__init__()
        self.db_mng = DBManager()

    def add_word(self):
        add_word_window = AddWord()
        add_word_window.get_add_word_window()

    def get_dictionary_window(self):
        #создание основного окна
        self.dict_root = tk.Tk()
        self.dict_root.title(self.window_name+' : RGSH Dictionary Editor')
        self.dict_root.iconbitmap(self.pic_path)
        self.dict_root.resizable(False, False)
        self.dict_root.attributes("-topmost", True)

        self.main_frame = tk.Frame(self.dict_root)
        self.manage_frame = ttk.LabelFrame(self.dict_root,
                                           borderwidth = 5, 
                                           relief = 'solid', 
                                           text = 'Управление')
        self.manage_frame.grid(row = 0, column = 0, 
                               pady = 5, padx = 5, 
                               sticky = 'nsew')
        self.create_button(ttk, self.manage_frame, 
                           'Добавить', self.add_word, 
                           20, 'normal', 0, 0)
        self.create_button(ttk, self.manage_frame, 
                           'Удалить', self.delete_row, 
                           20, 'normal', 0, 1)
        self.create_button(ttk, self.manage_frame, 
                           'Помощь', self.help_page, 
                           20, 'normal', 0, 2)

        self.find_frame = ttk.LabelFrame(self.dict_root,
                                         borderwidth=5,
                                         relief = 'solid', 
                                         text = 'Найти слово')
        self.find_frame.grid(row = 1, column = 0, 
                               pady = 5, padx = 5, 
                               sticky = 'nsew')
        self.find_entry = ttk.Entry(self.find_frame, width= 40)
        self.find_entry.grid(row=0, column=0,
                            pady = 5, padx = 5, 
                            sticky = 'nsew')
        self.create_button(ttk, self.find_frame, 
                           'Найти', self.find_word, 
                           24, 'normal', 0, 1)
        

        self.dict_frame = ttk.LabelFrame(self.dict_root,
                                         borderwidth=5,
                                         relief = 'solid', 
                                         text = 'Словарь')
        self.dict_frame.grid(row = 2, column = 0, 
                               pady = 5, padx = 5, 
                               sticky = 'nsew')
        
        #создание дерева словаря
        global dict_tree

        self.dict_tree_style = ttk.Style()
        self.dict_tree_style.configure("Treeview",
                background="white",
                foreground="black",
                rowheight=25,
                fieldbackground="white")
        self.dict_tree_style.configure("Treeview.Heading",
                background="lightgray",
                foreground="black",
                relief="raised")

        dict_tree = ttk.Treeview(self.dict_frame, style="Treeview")
        dict_tree.grid()

        dict_tree['columns'] = ('rus', 'en')
        dict_tree.column("#0", width=0, stretch=tk.NO)

        dict_tree.column("rus")
        dict_tree.heading("rus", text='Английский')

        dict_tree.column("en")
        dict_tree.heading("en", text='Русский')
        self.add_data_tree()
        self.alternate_colors()
        dict_tree.bind('<Button-1>', self.alternate_colors_click)


        self.dict_root.update_idletasks()
        w,h = self.get_center_window(self.dict_root)
        self.dict_root.geometry('+{}+{}'.format(w,h))
        self.dict_root.mainloop()

    def alternate_colors(self):
        for i, item in enumerate(dict_tree.get_children()):
            if i % 2 == 0:
                dict_tree.item(item, tags=('evenrow',))
            else:
                dict_tree.item(item, tags=('oddrow',))
        dict_tree.tag_configure('evenrow', background='white')
        dict_tree.tag_configure('oddrow', background='lightblue')

    def alternate_colors_click(self, event):
        for i, item in enumerate(dict_tree.get_children()):
            if i % 2 == 0:
                dict_tree.item(item, tags=('evenrow',))
            else:
                dict_tree.item(item, tags=('oddrow',))
        dict_tree.tag_configure('evenrow', background='white')
        dict_tree.tag_configure('oddrow', background='lightblue')

    def help_page(self):
        help_window = DictHelpWindow()
        help_window.get_help_window()

    def delete_row(self):
        selected_rows = dict_tree.selection()
        if selected_rows:
            for row in selected_rows:
                item = dict_tree.item(row)
                self.db_mng.delete_row(item["values"][0])
                dict_tree.delete(row)
                self.alternate_colors()

            get_info_msg('Выделенные строка или строки успешно удалены')
        else:
            get_error_msg('Для удаление данных выделите нужные строки')
            return

    def add_data_tree(self):
        dict_data = self.db_mng.get_dictionary()
        for key, value in dict_data.items():
            dict_tree.insert("","end",
                            text="", values=(key, value))
            self.alternate_colors()
            
    def find_word(self):
        value_to_find = self.find_entry.get()
        if value_to_find:
            dict_tree_rows = dict_tree.get_children()
            #сброс подсветок
            for item in dict_tree_rows:
                dict_tree.item(item, tags=())

            #поиск строки и подсветка
            for item in dict_tree_rows:
                values = dict_tree.item(item, 'values')
                if value_to_find in values:
                    dict_tree.item(item, tags = ('highlight'))
                    dict_tree.see(item) #прокрутка до объекта
                    break
            dict_tree.tag_configure('highlight', background='yellow')
            self.find_entry.delete(0, tk.END)
        else:
            get_error_msg('Введите слово для поиска')

class AddWord(Window):
    def __init__(self):
        super().__init__()
        self.db_mng = DBManager()

    def get_add_word_window(self):
        #создание основного окна
        self.add_word_root = tk.Tk()
        self.add_word_root.title(self.window_name)
        self.add_word_root.iconbitmap(self.pic_path)
        self.add_word_root.resizable(False, False)
        self.add_word_root.attributes("-topmost", True)

        #надпись Русский
        self.rus_label = tk.Label(self.add_word_root, 
                                  text='Английский',)
        self.rus_label.grid(row=0, column=0,
                            pady = 5, padx = 5, 
                            sticky = 'nsew')
        
        #надпись Английский
        self.en_label = tk.Label(self.add_word_root,
                                 text='Русский')
        self.en_label.grid(row=0, column=1,
                            pady = 5, padx = 5, 
                            sticky = 'nsew')

        #ввод русского значения
        self.rus_entry = ttk.Entry(self.add_word_root)
        self.rus_entry.grid(row=1, column=0,
                            pady = 5, padx = 5, 
                            sticky = 'nsew')
        
        #ввод английского значения
        self.en_entry = ttk.Entry(self.add_word_root)
        self.en_entry.grid(row=1, column=1,
                            pady = 5, padx = 5, 
                            sticky = 'nsew')
        
        #кнопка добавления пары слов
        self.add_button = ttk.Button(master=self.add_word_root,
                                     text='Добавить', 
                                     command=self.add_word)
        self.add_button.grid(row=2, columnspan=2,
                            pady = 5, padx = 5, 
                            sticky = 'nsew')

        #установка окна в центр и цикл
        self.add_word_root.update_idletasks()
        w,h = self.get_center_window(self.add_word_root)
        self.add_word_root.geometry('+{}+{}'.format(w,h))
        self.add_word_root.mainloop()


    def add_word(self):
        values_to_add = (self.rus_entry.get(), self.en_entry.get())
        for value in values_to_add:
            if value == '':
                get_error_msg(
                    'Введите данные в поля для ввода')
                return
            
        add_word_state = self.db_mng.add_word(values_to_add)
        if add_word_state:
            get_info_msg(
                'Пара слов успешно добавлена в базу')
            dict_tree.insert("", "end",
                             text='', values=(values_to_add[0], 
                                             values_to_add[1]))
            self.rus_entry.delete(0, tk.END)
            self.en_entry.delete(0, tk.END)
            return
        
        get_error_msg(
            'Одно из слов уже находится в базе.\nПроверьте правильность заполнения')

class DBManager:
    def __init__(self):
        self.folder_path = get_path()
        self.db_path = self.folder_path+"\\lib\\DICTIONARY.db"
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()
        self.table_name = 'dictionary'

    def create_table(self):
        self.cursor.execute("""
                            CREATE TABLE IF NOT EXISTS
                            {}
                            (id INTEGER PRIMARY KEY AUTOINCREMENT,
                            rus_word TEXT,
                            en_word TEXT)
                            """.format(self.table_name))

    def get_dictionary(self):
        self.cursor.execute("""
                            SELECT rus_word, en_word
                            FROM {}""".format(self.table_name))
        dict_list = self.cursor.fetchall()
        keys, values = zip(*dict_list)
        return dict(zip(keys, values))
            
    def get_column_info(self, column_name):
        self.cursor.execute("""SELECT {}
                            FROM {}""".format(column_name, self.table_name))
        column_info = self.cursor.fetchall()
        info_list = [info[0] for info in column_info]

        return info_list

    def add_word(self, values):
        dictionary = self.get_dictionary()
        for value in values:
            if value in dictionary:
                return False
            if value in dictionary.values():
                return False

        self.cursor.execute("""
                            INSERT INTO {}
                            (rus_word, en_word)
                            VALUES (?,?)
                               """.format(self.table_name), (values))
        self.conn.commit()

        return True
    
    def delete_row(self, value):
        delete_query = """DELETE from {} 
                        where rus_word = ?""".format(self.table_name)
        self.cursor.execute(delete_query, (value, ))
        self.conn.commit()

class DictHelpWindow(Window):
    def __init__(self):
        super().__init__()

    def get_help_window(self):
        help_root = tk.Tk()
        help_root.title(self.window_name+' : Справка')
        help_root.iconbitmap(self.pic_path)
        help_root.resizable(False, False)
        help_root.attributes("-topmost", True)
        
        help_frame = ttk.LabelFrame(help_root, borderwidth = 5, relief = 'solid', text = 'Справка')
        help_frame.grid(row = 0, column = 0, pady = 5, padx = 5, sticky = 'nsew')
        
        HelpPages = read_json('\\lib\\HELPPAGES.json' )
        HelpText = ''
        for item, text in HelpPages.items():
            if item == 'DICTIONARY EDITOR HELP':
                for stroka in text:
                    HelpText = HelpText+stroka+'\n'

        help_label = tk.Label(help_frame, text = HelpText, state=["normal"])
        help_label.grid(row = 0, column = 0, sticky = 'e', pady = 2, padx = 5)

        help_root.update_idletasks()
        w,h = self.get_center_window(help_root)
        help_root.geometry('+{}+{}'.format(w,h))
        help_root.mainloop()

def get_error_msg(error_msg):
    showerror('Ошибка', error_msg)

def get_info_msg(info_msg):
    showinfo('Информация', info_msg)

#dictio = DictionaryWindow()
#dictio.get_dictionary_window()

