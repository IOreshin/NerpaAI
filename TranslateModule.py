# -*- coding: utf-8 -*-
from NerpaUtility import KompasAPI
from tkinter import filedialog

import pythoncom
import win32com.client

class TranslateCDW(KompasAPI):
    def __init__(self):
        super().__init__()
        self.rus_paths = []
        self.iDocuments = self.app.Documents

    def get_cdw_docs(self):
        filepaths = filedialog.askopenfilenames(
            title = "Выбор чертежей", 
            filetypes = [("КОМПАС-Чертежи", "*.cdw")])
        if filepaths:
            for drawing_path in filepaths:
                self.resave_docs(drawing_path)
            
            for rus_path in self.rus_paths:
                self.translate_cdw(rus_path)


    def resave_docs(self, drawing_path):
        self.iDoc = self.iDocuments.Open(drawing_path, False, True)
        self.rus_filename = drawing_path[:-4]+' RUS.cdw'
        self.rus_paths.append(self.rus_filename)
        self.iDoc.SaveAs(self.rus_filename)

    def translate_cdw(self, rus_path):
        self.iDoc = self.iDocuments.Open(rus_path, True, False)
        self.iDoc2D = self.module.IKompasDocument2D(self.iDoc)
        self.iDoc2D1 = self.module.IKompasDocument2D1(self.iDoc2D)
        self.ViewsMng = self.iDoc2D.ViewsAndLayersManager
        self.iViews = self.ViewsMng.Views
        self.views_objects = []

        for i in range(1500):
            self.iView = self.iViews.View(float(i))
            if self.iView:
                
                safe_array = win32com.client.VARIANT(
                    pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH,(self.iView,))
                #print(safe_array)
                afa = self.iDoc2D1.DestroyObjects(safe_array)
                self.iView.Update()

                print(afa)

        
            
            



test = TranslateCDW()
test.get_cdw_docs()






