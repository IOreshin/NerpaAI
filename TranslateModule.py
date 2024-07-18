# -*- coding: utf-8 -*-
from NerpaUtility import KompasAPI, get_path
from tkinter import filedialog
import shutil
import win32com.client
from DictionaryModule import DBManager

from tkinter.messagebox import showerror, showinfo

class TranslateCDW(KompasAPI):
    def __init__(self):
        super().__init__()
        self.rus_paths = []
        self.iDocuments = self.app.Documents
        self.dict_path = get_path()+'\\lib\\DICTIONARY.csv'
        self.DICTIONARY = self.get_dictionary()
        self.dimensions_interfaces = ('DiametralDimensions',
                                      'LineDimensions',
                                      'AngleDimensions',
                                      'RadialDimensions',
                                      )
        self.dimension_interface = ('DiametralDimension',
                                    'LineDimension',
                                    'AngleDimension',
                                    'RadialDimension',
                                    )
        self.marking_interfaces = ('DrawingTexts',
                                   )
        self.marking_interface = ('DrawingText',
                                  )

    def get_dictionary(self):
        db_mng = DBManager()
        dictionary = db_mng.get_dictionary()
        db_mng.conn.close()
        return dictionary
    
    def get_cdw_docs(self):
        filepaths = filedialog.askopenfilenames(
            title = "Выбор чертежей", 
            filetypes = [("КОМПАС-Чертежи", "*.cdw")])
        if filepaths:
            save_state = self.resave_docs(filepaths)
            if save_state:
                for rus_path in self.rus_paths:
                    rus_doc = self.destroy_views(rus_path)
                    self.get_symbols_containers(rus_doc)
                    self.get_drawing_container(rus_doc)

                    rus_doc.Close(1)
            
                showinfo('Информация',
                     'Выбранные файлы созданы в версии RUS и переведены на русский язык. Проверьте и докорректируйте чертежи')
            
    def resave_docs(self, filepaths):
        try:
            for drawing_path in filepaths:
                drawing_path_rus = drawing_path[:-4]+' RUS.cdw'
                self.rus_paths.append(drawing_path_rus)
                shutil.copyfile(drawing_path, drawing_path_rus)
            return True
        except:
            showerror('Ошибка', 'Ошибка копирования выбранных компонентов')
            return False

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

    def get_drawing_container(self, doc_dispatch):
        views = self.get_views_collection(doc_dispatch)
        for view in views:
            iDrawingContainer = self.module.IDrawingContainer(view)
            for i, interface in enumerate(self.marking_interfaces):
                marking_interface = getattr(iDrawingContainer, interface)
                if marking_interface:
                    for n in range(len(marking_interface)):
                        marking_item = getattr(marking_interface, self.marking_interface[i])(n)
                        marking_item_text = self.module.IText(marking_item)
                        if marking_item_text.Str:
                            edited_text = self.edit_mark_str(marking_item_text.Str)
                            marking_item_text.Str = edited_text

                        marking_item.Update()

    def get_symbols_containers(self, doc_dispatch):
        views = self.get_views_collection(doc_dispatch)
        for view in views:
            self.get_container_operations('ISymbols2DContainer', view)

    def get_container_operations(self, container_name, view):
        container_dispatch = getattr(self.module, container_name)(view)
        for i, interface in enumerate(self.dimensions_interfaces): #для коллекции обозначений
            dim_interface = getattr(container_dispatch, interface) #получить интерфейс коллекции (DiamDimensions, LineDimensions)
            if dim_interface: #если он есть
                #для конкретного интерфейса (DiamDimension, LineDimension)
                for n in range(len(dim_interface)):
                    dim_item = getattr(dim_interface, self.dimension_interface[i])(n)
                    diam_dim_text = self.module.IDimensionText(dim_item)
                    if diam_dim_text.Prefix.Str:
                        edited_text = self.edit_symbol_str(diam_dim_text.Prefix.Str)
                        diam_dim_text.Prefix.Str = edited_text
                    if diam_dim_text.Suffix.Str:
                        edited_text = self.edit_symbol_str(diam_dim_text.Suffix.Str)
                        diam_dim_text.Suffix.Str = edited_text
                    if diam_dim_text.TextUnder.Str:
                        edited_text = self.edit_symbol_str(diam_dim_text.TextUnder.Str)
                        diam_dim_text.TextUnder.Str = edited_text

                    dim_item.Update()

    def edit_mark_str(self, str_to_edit: str):
        if str_to_edit.strip() in self.DICTIONARY.keys():
            edited_text = self.DICTIONARY[str_to_edit.strip()]
            if str_to_edit.startswith(' '):
                return ' '+edited_text
            if str_to_edit.endswith(' '):
                return edited_text+' '

            return edited_text
        
        if '\n' in str_to_edit:
            edited_text = ''
            for row in str_to_edit.split('\n'):
                print(row)
                edited_text += self.edit_single_str(row)
                edited_text += '\n'

        return edited_text
            

    def edit_symbol_str(self, str_to_edit: str):
        if str_to_edit.strip() in self.DICTIONARY.keys():
            edited_text = self.DICTIONARY[str_to_edit.strip()]
            if str_to_edit.startswith(' '):
                return ' '+edited_text
            if str_to_edit.endswith(' '):
                return edited_text+' '

            return edited_text
        
        edited_text = self.edit_single_str(str_to_edit)

        return edited_text
    
    def edit_single_str(self, str_to_edit:str):
        split_text = str_to_edit.split(' ')
        edited_text = ''
        for split_item in split_text:
            if split_item == 'PLACES':
                edited_text = self.get_correct_form(split_text[0], split_text[1])

            elif split_item in self.DICTIONARY.keys():
                edited_text += self.DICTIONARY[split_item]
            else:
                edited_text += split_item

            if split_text.index(split_item) != len(split_text)-1:
                edited_text += ' '

        return edited_text

    def get_correct_form(self, initial_number:str, second_try:str):
        try:
            number = int(initial_number)
            self.number_space = False
        except:
            self.number_space = True
            number = int(second_try)
        finally:
            if 11 <= number % 100 <= 14:
                if self.number_space:
                    return f" {number} МЕСТ"
                return f"{number} МЕСТ"
            
            elif number % 10 == 1:
                if self.number_space:
                    return f" {number} МЕСТО"
                return f"{number} МЕСТО"
            
            elif 2 <= number % 10 <= 4:
                if self.number_space:
                    return f" {number} МЕСТА"
                return f" {number} МЕСТА"
            else:
                if self.number_space:
                    return f" {number} МЕСТ"
                return f"{number} МЕСТ"
 

test = TranslateCDW()
test.get_cdw_docs()







