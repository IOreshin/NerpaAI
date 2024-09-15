# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Назначение модуля: для труб, полученных холодной гибкой,
# необходимо на чертеже приложить BEND TABLE с информацией всех углов:
# координаты X, Y, Z, BEND ANGLE, BEND TWIST, OFFSET.
# В этом модуле реализуется такая возможность:
# создается пользовательское окно, а также реализуется функция получения и
# обработки информации по линиям, а также запись в чертеж.
# Также здесь реализована запись TEMP файла. Это сделано специально,
# для упрощения работы модуля
#-------------------------------------------------------------------------------
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

from .WindowModule import Window
from .NerpaUtility import KompasAPI

class PDFWindow(Window): #класс для создания пользовательского окна
    def __init__(self):
        super().__init__()
        self.kAPI = KompasAPI()
        self.app = self.kAPI.app
        self.module = self.kAPI.module
        self.iConverter = self.app.Converter('C:\\Program Files\\ASCON\\KOMPAS-3D v22\\Bin\\Pdf2d.dll')

    def getWindow(self):
        root = tk.Tk()
        root.title(self.window_name)
        root.iconbitmap(self.pic_path)
        root.resizable(False, False)
        root.attributes("-topmost", True)

        frame1 = ttk.LabelFrame(root, borderwidth = 5, relief = 'solid', text = '')
        frame1.grid(row = 0, column = 0, pady = 5, padx = 5, sticky = 'nsew')

        def crtPDF():
            doc = self.app.ActiveDocument
            if doc.DocumentType in (1,3): #если чертеж, фрагмент или специя
                newName = doc.PathName
                snewName = newName[:-4]
                # IConverter.VisualEditConvertParam() #ничерта не работает
                self.iConverter.Convert(newName, snewName+'.pdf', 0, True)
                self.app.MessageBoxEx("Создан файл\n" + snewName+'.pdf', "PDF Module", 64)
            else:
                self.app.MessageBoxEx("Активный документ не является чертежом, фрагментом или спецификацией!", "PDF Module", 48)


        def crtmanyPDF():
            filepaths = filedialog.askopenfilenames(title = "Выбор чертежей", 
                                                    filetypes = [("КОМПАС-Чертежи", "*.cdw")])
            substrings = []
            in_brackets = False
            current_substring = ""
            for c in filepaths:
                if c == "{":
                    in_brackets = True
                elif c == "}" and in_brackets:
                    substrings.append(current_substring)
                    current_substring = ""
                    in_brackets = False
                elif in_brackets:
                    current_substring += c
            if current_substring:
                substrings.append(current_substring)
            for newName in substrings:
                snewName = newName[:-4]
                self.iConverter.Convert(newName, snewName+'.pdf', 0, True)


        def active_cdws_pdf():
            iDocuments = self.app.Documents
            for i in range(iDocuments.Count):
                iKompasDocument = iDocuments.Item(i)
                if iKompasDocument.DocumentType == 1: #чертеж
                    try:
                        pdf_name = ''.join([iKompasDocument.PathName[:-4], '.pdf'])
                        self.iConverter.Convert(iKompasDocument.PathName, pdf_name, 0, True)
                        self.app.MessageBoxEx("Создан файл\n" + pdf_name, "PDF Module", 64)
                    except Exception as e:
                        self.app.MessageBoxEx("Произошла ошибка при сохранении файла {}: {}".format(pdf_name, e))


        Window.create_button(self, ttk, frame1, 'ТЕКУЩИЙ ЧЕРТЕЖ В PDF', crtPDF, 40, 'normal', 0, 0)
        Window.create_button(self, ttk, frame1, 'ОТКРЫТЫЕ ЧЕРТЕЖИ В PDF', active_cdws_pdf, 40, 'normal', 1, 0)
        Window.create_button(self, ttk, frame1, 'ВЫБРАТЬ ЧЕРТЕЖИ ИЗ ПАПКИ', crtmanyPDF, 40, 'normal', 2, 0)

        # root.update_idletasks()
        # w,h = Window.get_center(root)
        # root.geometry('+{}+{}'.format(w,h))
        root.mainloop()