# -*- coding: utf-8 -*-

from .NerpaUtility import KompasAPI, get_path, get_resource_path

class PropertyManager(KompasAPI):
    '''
    Класс для удобного использования методов,
    связанных с работой со свойствами
    '''
    def __init__(self):
        super().__init__()
        self.iPropertyMng = self.module.IPropertyMng(self.app)

    def get_doc_properties(self):
        '''
        Функция получения списка свойств документа
        '''
        try:
            self.kompas_doc = self.app.ActiveDocument
            if self.kompas_doc.DocumentType not in [4,5]:
                self.app.MessageBoxEx('Ошибка формата', 
                                    'Ошибка формата', 64)
        except Exception:
            self.app.MessageBoxEx('Ошибка формата', 
                                    'Ошибка формата', 64)
            return
        
        properties = self.iPropertyMng.GetProperties(self.kompas_doc)
        properties_names = []
        for prop in properties:
            properties_names.append(prop.Name)
        return properties_names
    
    def check_add_properties(self):
        '''
        Функция для проверки и добавления свойств RGSH
        '''
        try:
            lib_path = get_resource_path('resources/lib/PROPERTIES.lpt')
            lib_properties = self.iPropertyMng.GetProperties(lib_path)
            doc_properties = self.get_doc_properties()
            for lib_prop in lib_properties:
                if lib_prop.Name not in doc_properties:
                    self.iPropertyMng.AddProperty(self.kompas_doc, lib_prop)

        except Exception:
            self.app.MessageBoxEx('Ошибка в пути к библиотеке свойств',
                                  'Ошибка', 64)
            return



