# -*- coding: utf-8 -*-
from NerpaUtility import KompasAPI

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog


from WindowModule import Window

class PDFWindow(Window): #класс для создания пользовательского окна
    def __init__(self):
        super().__init__() 
        self.kAPI = KompasAPI()
        self.app = self.kAPI.app
        self.module = self.kAPI.module

    def getWindow(self):
        root = tk.Tk()
        root.title(self.window_name+' : PDF CREATOR')
        root.iconbitmap(self.pic_path)
        root.resizable(False, False)
        root.attributes("-topmost", True)
        # root.geometry("150x100") 

        frame1 = ttk.LabelFrame(root, borderwidth = 5, relief = 'solid', text = '')
        frame1.grid(row = 0, column = 0, pady = 5, padx = 5, sticky = 'nsew')

        def crtPDF():
            doc = self.app.ActiveDocument
            if doc.DocumentType in (1,3): #если чертеж, фрагмент или специя
                newName = doc.PathName
                snewName = newName[:-4]
                IConverter = self.app.Converter('C:\\Program Files\\ASCON\\KOMPAS-3D v22\\Bin\\Pdf2d.dll')
                # IConverter.VisualEditConvertParam() 
                IConverter.Convert(newName, snewName+'.pdf', 0, True)
                self.app.MessageBoxEx("Создан файл\n" + snewName+'.pdf', "PDF Module", 64)
            else:
                self.app.MessageBoxEx("Активный документ не является чертежом, фрагментом или спецификацией!", "PDF Module", 48)

        def crtmanyPDF():
            filepaths = filedialog.askopenfilenames(title = "Выбор чертежей",
                                                    filetypes = [("КОМПАС-Чертежи", "*.cdw")])
            IConverter = self.app.Converter('C:\\Program Files\\ASCON\\KOMPAS-3D v22\\Bin\\Pdf2d.dll')
            for drawing in filepaths:
                newName = drawing
                snewName = newName[:-4]
                IConverter.Convert(newName, snewName+'.pdf', 0, True)

        self.create_button(ttk, frame1, 
                             'SAME FOLDER PDF', crtPDF, 
                             40, 'normal', 0, 0)
        self.create_button(ttk, frame1, 
                             'MANY PDF', crtmanyPDF, 
                             40, 'normal', 1, 0)

        # root.update_idletasks()
        # w,h = Window.get_center(root)
        # root.geometry('+{}+{}'.format(w,h))
        root.mainloop()