# -*- coding: utf-8 -*-

from .WindowModule import Window
from .NerpaUtility import KompasAPI, read_json
from .ConstantsRGSH import SHEET_FORMATS as formats
import tkinter as tk
from tkinter import ttk


class TechDemandWindow(Window):
    def __init__(self):
        super().__init__()
        self.kAPI = KompasAPI()
        self.get_tt_window()

    def get_tt_window(self):
        root = tk.Tk()
        root.title(self.window_name)
        root.iconbitmap(self.pic_path)
        root.attributes("-topmost", True)

        frame1 = ttk.LabelFrame(root, borderwidth = 5, relief = 'solid', text = 'Справка')
        frame1.grid(row = 0, column = 0, columnspan = 2, pady = 5, padx = 5, sticky = 'nsew')

        frame2 = ttk.LabelFrame(root, borderwidth = 5, relief = 'solid', text = '')
        frame2.grid(row = 1, column = 0, columnspan = 2, pady = 5, padx = 5, sticky = 'nsew')

        help_label = tk.Label(frame1, text = "Выберите из списка шаблон ТТ и нажмите кнопку <Добавить> ", state=["normal"])
        help_label.grid(row = 0, column = 1, sticky = 'e', pady = 2, padx = 5)

        self.templates = read_json('resources/lib/TECHDEMANDS.JSON')

        templates_names = []
        for name, text in self.templates.items():
            templates_names.append(name)

        combobox_type = ttk.Combobox(frame2, values=templates_names, width=40)
        combobox_type.current(0)
        combobox_type.grid(row = 0, column= 0, columnspan=2,padx=5,pady=5)

        def get_tt():
            try:
                doc = self.kAPI.app.ActiveDocument
                if doc.DocumentType != 1:
                    self.kAPI.app.MessageBoxEx('Активный документ не является чертежом',
                                            'Ошибка формата', 64)
                    return
            except Exception:
                self.kAPI.app.MessageBoxEx('Активный документ не является чертежом',
                                            'Ошибка формата', 64)
                return

            iKompasDocument2D = self.kAPI.module.IKompasDocument2D(doc)
            iDrawingDocument = self.kAPI.module.IDrawingDocument(iKompasDocument2D)
            TechDemand = iDrawingDocument.TechnicalDemand
            if TechDemand.IsCreated is True:
                self.kAPI.app.MessageBoxEx('Технические требования уже размещены на чертеже',
                                           'Ошибка', 64)
                return
            TechDemand.AutoPlacement = False
            TechText = TechDemand.Text
            for name, text in self.templates.items():
                if name == combobox_type.get():
                    TextLine = TechText.Add()
                    TextLine.Align = 1
                    TextItem = TextLine.Add()
                    TextItem.Str = 'NOTES:'
                    TextFont = self.kAPI.module.ITextFont(TextItem)
                    TextFont.Underline = True
                    TechDemand.Update()
                    for item in text:
                        TextLine = TechText.Add()
                        TextLine.Numbering = 1
                        TextLine.Str = item

            iLayoutSheets = doc.LayoutSheets
            iLayoutSheet = iLayoutSheets.Item(0)
            iSheetFormat = iLayoutSheet.Format
            Format = iSheetFormat.Format
            TechDemand.BlocksGabarits = ((formats[Format][1]-140), 65, formats[Format][1]-5, formats[Format][2])
            TechDemand.Update()
            self.kAPI.app.MessageBoxEx('Технические требования размещены на чертеже. Проверьте корректность записи',
                                       'Успех', 64)
            
        self.create_button(ttk, frame2, 'Добавить', get_tt, 15, 'normal', 0, 2)

        root.update_idletasks()
        w,h = self.get_center_window(root)
        root.geometry('+{}+{}'.format(w,h))
        root.mainloop()





