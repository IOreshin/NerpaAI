# -*- coding: utf-8 -*-
from .NerpaUtility import KompasAPI
import tkinter as tk
from tkinter import filedialog
import shutil
from .ConstantsRGSH import SHEET_FORMATS as formats
from .DictionaryModule import DBManager

import time
import os

from tkinter.messagebox import showerror, showinfo

class TranslateCDW(KompasAPI):
    def __init__(self):
        super().__init__()
        self.temp_dir = 'C:\\NerpaTranslateTemp'

        self.rus_paths = []
        self.iDocuments = self.app.Documents
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
        
        self.translate_cdw_docs()

    def get_dictionary(self):
        '''
        Получить словарь РГШ
        '''
        db_mng = DBManager()
        dictionary = db_mng.get_dictionary()
        db_mng.conn.close()
        return dictionary
    
    def translate_cdw_docs(self):
        '''
        Основная функция модуля. Запускает окно выбора файлов.
        Если выбраны документы, то запускает остальные методы
        для перевода чертежей на русский
        '''
        root = tk.Tk()
        root.withdraw()

        filepaths = filedialog.askopenfilenames(
            title = "Выбор чертежей", 
            filetypes = [("КОМПАС-Чертежи", "*.cdw")])
        file_list = root.tk.splitlist(filepaths)
        
        if file_list:
            #try:
               # os.mkdir(self.temp_dir)
            #except FileExistsError:
             #   pass


            start_time = time.time()
            save_state = self.resave_docs(file_list)
            if save_state:
                original_dir_items = file_list[0].split('/')
                self.original_dir_path = '\\'.join(original_dir_item 
                                              for original_dir_item in original_dir_items[:-1])
                
                for rus_path in self.rus_paths:
                    rus_doc = self.open_doc_and_destroy_views(rus_path)
                    self.get_drawing_operations(rus_doc)

                    rus_doc.Close(1)
                end_time = time.time()
                exuctution_time = round(end_time-start_time,2)
                
                #if self.copy_to_original_dir():
                showinfo('Информация',
                    'Выбранные файлы созданы в версии RUS и переведены на русский язык.Проверьте и докорректируйте чертежи.Перевод выполнен за {} сек'
                    .format(exuctution_time))

               # else:
                #    showerror('Ошибка',
                #             'Во время копирования в {} произошла ошибка'
                 #            .format(self.original_dir_path))

    def resave_docs(self, filepaths):
        '''
        Метод копирования исходных файлов.
        Дополнительно записывает пути файлов в 
        drawing_path_rus
        '''
        try:
            for drawing_path in filepaths:
                drawing_path_rus = drawing_path[:-4]+' RUS.cdw'
                self.rus_paths.append(drawing_path_rus)
                shutil.copyfile(drawing_path, drawing_path_rus)
            return True
        except:
            showerror('Ошибка', 'Ошибка копирования выбранных компонентов')
            return False

    def resave_docs_(self, filepaths):
        '''
        Метод копирования исходных файлов.
        Дополнительно записывает пути файлов в 
        drawing_path_rus
        '''
        try:
            for drawing_path in filepaths:
                doc_name_items = drawing_path.split('/')
                doc_name = doc_name_items[-1]
                name, extension = os.path.splitext(doc_name)
                path_items = [self.temp_dir,'\\',name, ' RUS.cdw']
                drawing_path_rus = ''.join(path_items)
                self.rus_paths.append(drawing_path_rus)
                shutil.copyfile(drawing_path, drawing_path_rus)
            return True
        except:
            showerror('Ошибка', 'Ошибка копирования выбранных компонентов')
            return False

    def copy_to_original_dir(self):
        #try:
            for rus_path in self.rus_paths:
                doc_name_items = rus_path.split('\\')
                doc_name = '\\'+doc_name_items[-1]
                print(rus_path)
                print(self.original_dir_path+doc_name)
                shutil.copyfile(rus_path, self.original_dir_path+doc_name)
            

            shutil.rmtree(self.temp_dir)
            #os.rmdir(self.temp_path)
            return True
        
        #except Exception:
        #    return False
  
    def get_views_collection(self, doc_dispatch):
        '''
        Метод получения всех возможных видов на чертеже.
        Возвращает список диспатчей на каждый вид
        '''
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

    def open_doc_and_destroy_views(self, rus_path):
        '''
        Функция разрушения всех видов на чертеже
        '''
        iDoc = self.iDocuments.Open(rus_path, False, False)
        iDoc2D = self.module.IKompasDocument2D(iDoc)
        iDoc2D1 = self.module.IKompasDocument2D1(iDoc2D)
        views = self.get_views_collection(iDoc)
        for view in views:
            iDoc2D1.DestroyObjects(view) 

        return iDoc

    def get_drawing_operations(self, doc_dispatch):
        '''
        Функция для запуска прогона перевода
        '''
        views = self.get_views_collection(doc_dispatch)
        for view in views:
            self.get_container_operations('ISymbols2DContainer', view)
            self.drawing_container_operations(view)
            self.translate_bom_table(view, doc_dispatch)
            self.translate_drawing_tables(view)

        self.translate_stamp(doc_dispatch)
        self.translate_tech_demands(doc_dispatch)
        #self.change_layout_sheets(doc_dispatch)

    def destroy_object(self, doc_dispatch, object):
        '''
        Общий метод для удаления различных объектов на чертеже
        '''
        iDoc2D = self.module.IKompasDocument2D(doc_dispatch)
        iDoc2D1 = self.module.IKompasDocument2D1(iDoc2D)
        iDoc2D1.DestroyObjects(object)

    def translate_bom_table(self, view, doc_dispatch):
        '''
        Метод для переименования компонентов BOM
        '''
        if view.Name in ['BOM']:
            iSymbolsContainer = self.module.ISymbols2DContainer(view)
            DrwTables = iSymbolsContainer.DrawingTables
            for i in range(DrwTables.Count):
                DrawingTable = DrwTables.DrawingTable(i)
                table = self.module.ITable(DrawingTable)
                table_object = self.module.IDrawingObject(table)
                if table.RowsCount == 1: #шапка
                    cell = table.Cell(0,0)
                    cell_text = self.module.IText(cell.Text)
                    cell_text.Str = 'СОСТАВ ИЗДЕЛИЯ'
                    table_object.Update()

            AssoTables = iSymbolsContainer.AssociationTables
            BomTable = AssoTables.AssociationTable(0)
            self.destroy_object(doc_dispatch, BomTable)

            DrwTables = iSymbolsContainer.DrawingTables
            table = self.module.ITable(DrwTables(1))
            self.translate_any_table(table, True)

    def translate_drawing_tables(self, view):
        if view.Name not in ('BOM'):
            iSymbolsContainer = self.module.ISymbols2DContainer(view)
            DrwTables = iSymbolsContainer.DrawingTables
            for i in range(DrwTables.Count):
                DrawingTable = DrwTables.DrawingTable(i)
                table = self.module.ITable(DrawingTable)
                if table.RowsCount != 1:
                    self.translate_any_table(table, True)

    def translate_any_table(self, table, hyper_text_state:bool):
        table_object = self.module.IDrawingObject(table)
        columns_count = table.ColumnsCount
        rows_count = table.RowsCount
        table_range = table.Range(0,0,rows_count,columns_count)
        cells = table_range.Cells
        for cell in cells: 
            cell_format = self.module.ICellFormat(cell)
            cell_format.ReadOnly = False
            table_object.Update()
            cell_text = self.module.IText(cell.Text)
            if cell_text.Str and cell_text.Count == 1:
                iTextLine = cell_text.TextLine(0)
                for i in range(iTextLine.Count):
                    iTextItem = iTextLine.TextItem(i)
                    if hyper_text_state:
                        HyperTextParams = self.module.IHypertextReferenceParam(iTextItem)
                        HyperTextParams.Destroy()
                        iTextItem.Update()
                    if iTextItem.Str:
                        edited_text = self.edit_symbol_str(iTextItem.Str)
                        iTextItem.Str = edited_text
                        iTextItem.Update()
            
            #случай нескольких строк в ячейке таблицы
            elif cell_text.Count > 1:
                TextLines = cell_text.TextLines
                full_en_text = []
                for i,TextLine in enumerate(TextLines):
                    for n in range(TextLine.Count):
                        TextItem = TextLine.TextItem(n)
                        full_en_text.append([TextItem.Str, i, n])

                en_words = [word[0] for word in full_en_text]
                en_str = " ".join(en_words)
                ru_str = self.edit_mark_str(en_str)
                ru_words = ru_str.split(' ')
                #случай совпадения количества слов в русской и английской версиях
                if len(en_words) == len(ru_words):
                    for i,rus_word in enumerate(ru_words):
                        TextLine = cell_text.TextLine(full_en_text[i][1])
                        TextItem = TextLine.TextItem(full_en_text[i][2])
                        TextItem.Str = rus_word
                        TextItem.Update()
                else:
                    cell_text.Clear()
                    for rus_word in ru_words:
                        TextLine = cell_text.Add()
                        TextItem = TextLine.Add()
                        TextItem.Str = rus_word
                        TextItem.Update()
            else:
                continue

    def drawing_container_operations(self, view):
        iDrawingContainer = self.module.IDrawingContainer(view)
        iDrawingTexts = iDrawingContainer.DrawingTexts
        if iDrawingTexts.Count:
            for i in range(iDrawingTexts.Count):
                iDrawingText = iDrawingTexts.DrawingText(i)
                iDrawingText_text = self.module.IText(iDrawingText)
                #случай, когда в названии вида вставлена ссылка
                if iDrawingText_text.Str.startswith('^'):
                    self.translate_view_designation(view)
                    continue

                if iDrawingText_text.Str and not iDrawingText_text.Str.startswith('^'):
                    edited_text = self.edit_mark_str(iDrawingText_text.Str)
                    iDrawingText_text.Str = edited_text

                iDrawingText.Update()

    def translate_view_designation(self, view):
        view_designation = self.module.IViewDesignation(view)
        view_inscription = view_designation.DrawingText
        view_text = self.module.IText(view_inscription)
        for i in range(view_text.Count):
            text_line = view_text.TextLine(i)
            if not text_line.Str.startswith('^'):
                edited_text = self.edit_symbol_str(text_line.Str)
                text_line.Str = edited_text
                view_inscription.Update()

            if text_line.Str.startswith('^'):
                print(text_line.Str)
                for n in range(text_line.Count):
                    text_item = text_line.TextItem(n)
                    text_font = self.module.ITextFont(text_item)
                    text_font.Height = 7.0
                    text_font.Italic = True
                    view_inscription.Update()
                    text_item.Update()
                    

            view_inscription.Update()

    def translate_stamp(self, doc_dispatch):
        iLayoutSheets = doc_dispatch.LayoutSheets
        for i in range(iLayoutSheets.Count):
            iLayoutSheet = iLayoutSheets.Item(i)
            iStamp = iLayoutSheet.Stamp
            text_cell_counter = 0
            while text_cell_counter < 1000:
                text = iStamp.Text(text_cell_counter)
                if text.Str:
                    edited_text = self.edit_mark_str(text.Str)
                    text.Str = edited_text
                    iStamp.Update()
                
                text_cell_counter += 1

    def translate_tech_demands(self, doc_dispatch):
        iDoc2D = self.module.IKompasDocument2D(doc_dispatch)
        iDrawingDocument = self.module.IDrawingDocument(iDoc2D)
        TechDemands = iDrawingDocument.TechnicalDemand
        if TechDemands.IsCreated is False:
            return
        
        key_words = ['NOTES',
                     'CONJUNCTIONS',
                     '3D MODEL',
                     'CONJUNCTION']
        TechText = TechDemands.Text
        TextLines = TechText.TextLines
        
        for TextLine in TextLines:
            if TextLine:
                if any(key_word in TextLine.Str for key_word in key_words): 
                    TextLine.Delete()
                    TechDemands.Update()

        TechText = TechDemands.Text
        TextLines = TechText.TextLines
        tech_paragraphs = []
        paragraph_count = -1
        for TextLine in TextLines:
            new_paragraph_state = False
            if TextLine.Numbering == 1:
                new_paragraph_state = True
            if new_paragraph_state:
                paragraph_count += 1
            try:
                tech_paragraphs[paragraph_count] = tech_paragraphs[
                paragraph_count]+' '+TextLine.Str
            except IndexError:
                tech_paragraphs.append(TextLine.Str)

        TechDemands.Delete()
        TechDemands.Update()
        for tech_paragraph in tech_paragraphs:
            edited_text = self.edit_mark_str(tech_paragraph)
            TextLine = TechText.Add()
            TextLine.Str = edited_text
            TextLine.Numbering = 1

        iLayoutSheets = doc_dispatch.LayoutSheets
        iLayoutSheet = iLayoutSheets.Item(0)
        iSheetFormat = iLayoutSheet.Format
        Format = iSheetFormat.Format
        TechDemands.BlocksGabarits = ((formats[Format][1]-140), 65, formats[Format][1]-5, formats[Format][2])
        TechDemands.Update()

    def change_layout_sheets(self, doc_dispatch):
        iLayoutSheets = doc_dispatch.LayoutSheets
        for i in range(iLayoutSheets.Count):
            iLayoutSheet = iLayoutSheets.Item(i)
            #iLayoutSheet.LayoutStyleNumber = 2
            iLayoutSheet.Update()

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
        #если целиком фраза обозначения есть в словаре
        if str_to_edit.strip() in self.DICTIONARY.keys():
            edited_text = self.DICTIONARY[str_to_edit.strip()]
            if str_to_edit.startswith(' '):
                return ' '+edited_text
            if str_to_edit.endswith(' '):
                return edited_text+' '
            return edited_text
        
        #если фраза состоит из строк
        if '\n' in str_to_edit:
            edited_text = ''
            for row in str_to_edit.split('\n'):
                edited_text += self.edit_single_str(row)
                edited_text += '\n'
            return edited_text

        return self.edit_symbol_str(str_to_edit)
            
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
                    return " {} МЕСТ".format(number)
                return "{} МЕСТ".format(number)
            
            elif number % 10 == 1:
                if self.number_space:
                    return " {} МЕСТО".format(number)
                return "{} МЕСТО".format(number)
            
            elif 2 <= number % 10 <= 4:
                if self.number_space:
                    return " {} МЕСТА".format(number)
                return " {} МЕСТА".format(number)
            else:
                if self.number_space:
                    return " {} МЕСТ".format(number)
                return "{} МЕСТ".format(number)

