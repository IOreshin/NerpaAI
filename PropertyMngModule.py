# -*- coding: utf-8 -*-

from NerpaUtility import KompasAPI, format_error, get_path
from tkinter.messagebox import showerror

class PropertyManager():
    def __init__(self):
        self.kAPI = KompasAPI()
        self.module = self.kAPI.module
        self.app = self.kAPI.app
        self.iPropertyMng = self.module.IPropertyMng(self.app)

    
    def get_doc_properties(self):
        self.kompas_doc = self.kAPI.app.ActiveDocument
        if self.kompas_doc.DocumentType not in [4,5]:
            format_error('ASSY OR DETAIL')
            return
        
        properties = self.iPropertyMng.GetProperties(self.kompas_doc)
        properties_names = []
        for prop in properties:
            properties_names.append(prop.Name)
        return properties_names
    
    def check_add_properties(self):
        try:
            directory_path = get_path()
            lib_path = directory_path+'\\lib\\PROPERTIES.lpt'
            lib_properties = self.iPropertyMng.GetProperties(lib_path)
            doc_properties = self.get_doc_properties()
            for lib_prop in lib_properties:
                if lib_prop.Name not in doc_properties:
                    self.iPropertyMng.AddProperty(self.kompas_doc, lib_prop)

        except Exception:
            showerror('ERROR', 'LIB PATH ERROR')
            return



