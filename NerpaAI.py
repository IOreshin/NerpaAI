# -*- coding: utf-8 -*-

from resources import *
import pythoncom
import tkinter as tk
from tkinter import ttk
import webbrowser
from PIL import Image, ImageTk
import threading

class SplashScreen(Window):
    def __init__(self, parent, on_close):
        super().__init__()
        self.window = tk.Toplevel(parent)
        self.window.title("NerpaAI")
        self.window.iconbitmap(self.pic_path)
 
        self.window.resizable(False, False)
        self.window.attributes("-topmost", True)  # Убедитесь, что окно поверх всех

        # Загрузка и отображение изображения PNG
        self.image_path = get_resource_path(
            'resources\\pic\\ICO RGSH PS 400X400.png')  # Укажите путь к вашему PNG файлу
        self.image = Image.open(self.image_path)
        self.photo = ImageTk.PhotoImage(self.image)

        self.label = tk.Label(self.window, image=self.photo)
        self.label.pack(expand=True, fill=tk.BOTH)
        #self.window.update_idletasks()
        w, h = self.get_center_window(self.window)
        self.window.geometry('+{}+{}'.format(w, h))

        # Запуск таймера для закрытия окна
        self.window.after(3000, self.close)
        self.on_close = on_close

    def close(self):
        self.window.destroy()
        self.on_close()

class LoadingWindow(Window):
    def __init__(self, parent):
        super().__init__()
        self.window = tk.Toplevel(parent)
        self.window.title("Ожидание")
        self.window.iconbitmap(self.pic_path)
        #self.window.geometry("200x100")
        self.window.resizable(False, False)
        self.window.grab_set()  # Захватываем фокус на это окно
        self.window.attributes("-topmost", True)  # Убедитесь, что окно поверх всех

        self.gif_path = get_resource_path('resources\\pic\\NerpaLoad.gif')  # Укажите путь к вашему GIF
        self.gif = Image.open(self.gif_path)
        self.frames = [ImageTk.PhotoImage(self.gif.copy().convert('RGBA'))]
        try:
            while True:
                self.gif.seek(len(self.frames))
                self.frames.append(ImageTk.PhotoImage(self.gif.copy().convert('RGBA')))
        except EOFError:
            pass

        #self.window.update_idletasks()
        w, h = self.get_center_window(self.window)
        self.window.geometry('+{}+{}'.format(w, h))

        self.label_image = tk.Label(self.window)
        self.label_image.pack()

        self.update_image(0)
        self.current_frame = 0
    
    def update_image(self, frame_number):
        self.label_image.config(image=self.frames[frame_number])
        self.current_frame = (frame_number + 1) % len(self.frames)
        self.window.after(25, self.update_image, self.current_frame)  # Обновление кадра каждые 100 мс

    def close(self):
        self.window.grab_release()  # Отпускаем захват фокуса
        self.window.destroy()

class MainWindow(Window):
    def __init__(self):
        super().__init__()
        self.app = NerpaUtility.KompasAPI().app
        self.root = tk.Tk()
        self.root.title(self.window_name)
        self.root.resizable(False, False)
        self.root.iconbitmap(self.pic_path)
        self.root.attributes("-topmost", True)
        self.root.withdraw() # Скрыть главное окно до завершения заставки

        splash_thread = threading.Thread(target=self.show_splash_screen)
        splash_thread.start()

        
    def show_splash_screen(self):
        splash = SplashScreen(self.root, self.show_main_window)
        

    def show_main_window(self):
        self.root.deiconify()  # Показать главное окно
        self.initialize_ui()
    

    def execute_with_loading(self, func, *args, **kwargs):
        def task():
            pythoncom.CoInitialize()
            try:
                func(*args, **kwargs)
            finally:
                pythoncom.CoUninitialize()
                self.root.after(0, loading_window.close)
        
        loading_window = LoadingWindow(self.root)  # Создание окна уведомления
        thread = threading.Thread(target=task).start()


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
             'command': lambda: self.execute_with_loading(AdaptAssy), 'state': 'normal'},
            {'text': 'Адаптировать активную деталь', 'frame': 'assembly', 
             'command': AdaltDetail, 'state': 'normal'},
            {'text': 'Добавить свойства ISO-RGSH', 'frame': 'assembly', 
             'command': self.check_add_prop, 'state': 'normal'},
            {'text': 'Добавить BOM', 'frame': 'drawing', 
             'command': BOMMaker, 'state': 'normal'},
            {'text': 'Создать МТО', 'frame': 'extra', 
             'command': MTOMaker, 'state': 'normal'},
            {'text': 'Установить позиции в активной сборке', 'frame': 'drawing', 
             'command': lambda: self.execute_with_loading(SetPositions), 'state': 'normal'},
            {'text': 'Добавить тех. требования', 'frame': 'drawing', 
             'command': TechDemandWindow, 'state': 'normal'},
            {'text': 'Добавить таблицу гибов', 'frame': 'drawing', 
             'command': BTWindow, 'state': 'normal'},
            {'text': 'Перевести .CDW чертежи ENG-RUS', 'frame': 'drawing',
              'command': TranslateCDW, 'state': 'normal'},
            {'text': 'Редактор словаря', 'frame': 'extra', 
             'command': DictionaryWindow, 'state': 'normal'},
            {'text': 'Создать PDF', 'frame': 'extra', 
             'command': self.get_pdf_window, 'state': 'normal'},
            {'text': 'Справка', 'frame': 'extra', 
             'command': self.get_help_page, 'state': 'normal'},
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

    def get_help_page(self):
        page_path = get_resource_path('ReadMe.htm')
        webbrowser.open(page_path)

if __name__ == '__main__':
    app = MainWindow()
    app.run()