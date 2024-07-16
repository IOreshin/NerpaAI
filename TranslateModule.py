# -*- coding: utf-8 -*-
from NerpaUtility import KompasAPI, get_path
from tkinter import filedialog
import shutil
import win32com.client

import csv


class TranslateCDW(KompasAPI):
    def __init__(self):
        super().__init__()
        self.rus_paths = []
        self.iDocuments = self.app.Documents
        self.dict_path = get_path()+'\\lib\\DICTIONARY.csv'
        self.DICTIONARY = self.get_dictionary()

    def get_dictionary_(self):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel_doc = excel.Workbooks.Open(self.dict_path)
        sheet = excel_doc.ActiveSheet
        last_row = sheet.Cells(sheet.Rows.Count, 1).End(
            -4162).Row
        data_dict = {}
        for row in range(2, last_row + 1):  # Начинаем считывать данные с 2 строки
            key = sheet.Cells(row, 1).Value
            value = sheet.Cells(row, 2).Value
            data_dict[key] = value
        excel_doc.Close(SaveChanges=False)
        excel.Quit()
        return data_dict
    
    def get_dictionary(self):
        data_dict = {}
        with open(self.dict_path, 'r', newline='', encoding='utf-8') as dict:
            c = csv.reader(dict)
            for row in c:
                key = row[0]
                value = row[1]
                data_dict[key] = value
                
        return data_dict
    
    def get_cdw_docs(self):
        filepaths = filedialog.askopenfilenames(
            title = "Выбор чертежей", 
            filetypes = [("КОМПАС-Чертежи", "*.cdw")])
        if filepaths:
            self.resave_docs(filepaths)
            for rus_path in self.rus_paths:
                rus_doc = self.destroy_views(rus_path)
                self.get_symbols_containers(rus_doc)

                rus_doc.Close(1)

    def resave_docs(self, filepaths):
        for drawing_path in filepaths:
            drawing_path_rus = drawing_path[:-4]+' RUS.cdw'
            self.rus_paths.append(drawing_path_rus)
            shutil.copyfile(drawing_path, drawing_path_rus)

    def get_views_collection(self, doc_dispatch):
        iDoc2D = self.module.IKompasDocument2D(doc_dispatch)
        ViewsMng = iDoc2D.ViewsAndLayersManager
        iViews = ViewsMng.Views
        #механизм обхода всех возможных видов,
        #нумерация самих видов может быть любой
        views_count = iViews.Count
        views_counter = 0
        views_dispatchs = []
        while views_counter < views_count:
            iView = iViews.View(views_counter)
            if iView:
                views_dispatchs.append(iView)
                views_counter += 1

        return views_dispatchs

    def destroy_views(self, rus_path):
        iDoc = self.iDocuments.Open(rus_path, False, False)
        iDoc2D = self.module.IKompasDocument2D(iDoc)
        iDoc2D1 = self.module.IKompasDocument2D1(iDoc2D)
        views = self.get_views_collection(iDoc)
        for view in views:
            iDoc2D1.DestroyObjects(view) 

        return iDoc

    def get_symbols_containers(self, doc_dispatch):
        views = self.get_views_collection(doc_dispatch)
        for view in views:
            iSymbols2DContainer = self.module.ISymbols2DContainer(view)
            #DIAMETRAL DIMENSIONS
            DiametralDimensions = iSymbols2DContainer.DiametralDimensions
            if DiametralDimensions:
                for i, dimension in enumerate(DiametralDimensions):
                    diam_dim = DiametralDimensions.DiametralDimension(i)
                    diam_dim_text = self.module.IDimensionText(diam_dim)
                    if diam_dim_text.Prefix.Str:
                        split_text = diam_dim_text.Prefix.Str.split(' ')
                        edited_text = ''
                        for split_item in split_text:
                            if split_item in self.DICTIONARY.keys():
                                edited_text += self.DICTIONARY[split_item]
                            else:
                                edited_text += split_item
                        edited_text += ' '
                        diam_dim_text.Prefix.Str = edited_text

                    diam_dim.Update()


    def get_correct_form(self, number):
        if 11 <= number % 100 <= 14:
            return f"{number} мест"
        elif number % 10 == 1:
            return f"{number} место"
        elif 2 <= number % 10 <= 4:
            return f"{number} места"
        else:
            return f"{number} мест"

test = TranslateCDW()
test.get_cdw_docs()







