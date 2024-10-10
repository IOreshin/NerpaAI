# -*- coding: utf-8 -*-
from .NerpaUtility import KompasAPI
import tkinter as tk
from tkinter import filedialog, messagebox
import shutil
from .ConstantsRGSH import SHEET_FORMATS as formats
from .DictionaryModule import DBManager

import time

from tkinter.messagebox import showerror

class TranslateCDW(KompasAPI):
    def __init__(self):
        super().__init__()

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
        db_mng = DBManager('\lib\DICTIONARY.db')
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

                if messagebox.askokcancel('Завершение работы',
                                          'Предварительный перевод чертежей выполнен за {} сек. Хотите открыть файлы?'.format(exuctution_time)):
                    for rus_path in self.rus_paths:
                        try:
                            self.iDocuments.Open(rus_path, True, False)
                        except Exception as e:
                            print(e)
                            continue

                return

    def resave_docs(self, filepaths):
        '''
        Метод копирования исходных файлов.
        Дополнительно записывает пути файлов в
        drawing_path_rus
        '''
        try:
            for drawing_path in filepaths:
                if 'RUS' not in drawing_path:
                    drawing_path_rus = drawing_path[:-4]+' RUS.cdw'
                    self.rus_paths.append(drawing_path_rus)
                    shutil.copyfile(drawing_path, drawing_path_rus)
            return True
        except:
            showerror('Ошибка', 'Ошибка копирования выбранных компонентов')
            return False

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
            if view.Name not in ['BEND TABLE TYP', 'BEND TABLE CS']:
                iBreakViewParam = self.module.IBreakViewParam(view)
                if iBreakViewParam.BreaksCount == 0:
                    iDoc2D1.DestroyObjects(view)

        return iDoc

    def get_drawing_operations(self, doc_dispatch):
        '''
        Функция для запуска прогона перевода
        '''
        views = self.get_views_collection(doc_dispatch)
        try:
            progress_bar = self.get_progress_bar()
            progress_bar.Start(0.0, len(views), '', True)
            for i, view in enumerate(views):
                progress_bar.SetProgress(i, '', True)
                if view.Name not in ['BEND TABLE TYP', 'BEND TABLE CS']:
                    self.get_container_operations('ISymbols2DContainer', view)
                    self.drawing_container_operations(view)
                    self.translate_bom_table(view, doc_dispatch)
                    self.translate_drawing_tables(view)
                else:
                    self.change_fragment(view, doc_dispatch)

            self.translate_stamp(doc_dispatch)
            self.translate_tech_demands(doc_dispatch)
            self.change_layout_sheets(doc_dispatch)
            progress_bar.Stop('', True)
        except Exception as e:
            print(e, view.Name, doc_dispatch.Name)
            progress_bar.Stop('', True)

    def change_fragment(self, view, doc_dispatch):
        doc_name = doc_dispatch.Name
        SDU_flag = False
        if 'SDU' in doc_name:
            SDU_flag = True

        DrawingContainer = self.module.IDrawingContainer(view)
        InsertionObjects = DrawingContainer.InsertionObjects
        InsertionObject = InsertionObjects.InsertionObject(0)
        if InsertionObject:
            InsertionDefinition = InsertionObject.InsertionDefinition
            if view.Name in ['BEND TABLE TYP']:
                InsertionDefinition.FileName = 'Z:\\14 РГШ\\YK RGS 2023\\02 Engineering\\08 MAN\\02 Engineering\\02 CAD folder\\WORKSHOP\\CONFIG KOMPAS\\SYMBOL\\BEND_TABLE_TYP_RUS.frw'
            if view.Name in ['BEND TABLE CS']:
                if SDU_flag:
                    InsertionDefinition.FileName = 'Z:\\14 РГШ\\YK RGS 2023\\02 Engineering\\08 MAN\\02 Engineering\\02 CAD folder\\WORKSHOP\\CONFIG KOMPAS\\SYMBOL\\BEND_TABLE_CS_SDU_RUS.frw'
                else:
                    InsertionDefinition.FileName = 'Z:\\14 РГШ\\YK RGS 2023\\02 Engineering\\08 MAN\\02 Engineering\\02 CAD folder\\WORKSHOP\\CONFIG KOMPAS\\SYMBOL\\BEND_TABLE_CS_RUS.frw'

            view.Update()

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
                for i, TextLine in enumerate(TextLines):
                    for n in range(TextLine.Count):
                        TextItem = TextLine.TextItem(n)
                        #добавление в массив английского выражения конструкции:
                        # [текст, строка, порядок в строке]
                        full_en_text.append([TextItem.Str, i, n])

                en_words = [word[0] for word in full_en_text]
                en_str = ' '.join(en_words)
                ru_str = self.edit_mark_str(en_str)

                ru_words = ru_str.split(' ')

                #если совпадает кол-во слов в русском и английском
                if len(en_words) == len(ru_words):
                    for i, rus_word in enumerate(ru_words):
                        TextLine = cell_text.TextLine(full_en_text[i][1])
                        TextItem = TextLine.TextItem(full_en_text[i][2])
                        TextItem.Str = rus_word
                        TextItem.Update()
                else:
                    cell_text.Clear()
                    TextLine = cell_text.Add()
                    for i,rus_word in enumerate(ru_words):
                        TextItem = TextLine.Add()
                        TextItem.Str = rus_word
                        if i != len(ru_words)-1:
                            TextItem = TextLine.Add()
                            TextItem.Str = ' '
                            TextItem.Update()
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

                if iDrawingText_text.Str and not iDrawingText_text.Str.startswith('^'):
                    view_name_flag = False
                    iTextLine = iDrawingText_text.TextLine(0)
                    for TextItem in iTextLine.TextItems:
                        if TextItem.Str.strip() in ['ISO VIEW', 'VIEW', 'TOP VIEW',
                                            'BOTTOM VIEW', 'ISO VIEW BOTTOM',
                                            'STB VIEW', 'FORE VIEW',
                                            'PORT VIEW', 'AFT VIEW',
                                            'ISO VIEW 1', 'ISO VIEW 2',
                                            'ISO VIEW 3', 'ISO VIEW 4',
                                            'ISO VIEW 5', 'ISO VIEW 6',
                                            'ISO VIEW 7', 'ISO VIEW 8']:
                            view_name_flag = True

                    if not view_name_flag:
                        edited_text = self.edit_mark_str(iDrawingText_text.Str)
                        iDrawingText_text.Str = edited_text
                        iDrawingText.Update()
                    if view_name_flag:
                        iTextLine = iDrawingText_text.TextLine(0)
                        for TextItem in iTextLine.TextItems:
                            iTextFont = self.module.ITextFont(TextItem)
                            iTextFont.Height = 7
                            TextItem.Update()
                            iDrawingText.Update()
                        if iDrawingText_text.Count > 1:
                            for i in range(iDrawingText_text.Count-1):
                                iTextLine = iDrawingText_text.TextLine(i+1)
                                if iTextLine:
                                    edited_text = self.edit_mark_str(iTextLine.Str)
                                    iTextLine.Str = edited_text
                                    iDrawingText.Update()

                if iDrawingText_text.Str.startswith('^'):
                    if iDrawingText_text.Count>1:
                        for i in range(iDrawingText_text.Count-1):
                            iTextLine = iDrawingText_text.TextLine(i+1)
                            if iTextLine:
                                edited_text = self.edit_mark_str(iTextLine.Str)
                                iTextLine.Str = edited_text

                        iDrawingText.Update()

                        iTextLine = iDrawingText_text.TextLine(0)
                        for TextItem in iTextLine.TextItems:
                            print(TextItem.Str)
                            iTextFont = self.module.ITextFont(TextItem)
                            iTextFont.Height = 7
                            TextItem.Update()

                        iDrawingText.Update()

    def translate_stamp(self, doc_dispatch):
        iLayoutSheets = doc_dispatch.LayoutSheets
        for i in range(iLayoutSheets.Count):
            iLayoutSheet = iLayoutSheets.Item(i)
            iStamp = iLayoutSheet.Stamp
            text_cell_counter = 0
            while text_cell_counter < 1000:
                text = iStamp.Text(text_cell_counter)
                if text.Str:
                    if text_cell_counter == 220:
                        text.Str = '13.09.2024'
                        iStamp.Update()
                    else:
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
        tube_bending_flag = False
        for tech_paragraph in tech_paragraphs:
            edited_text = self.edit_mark_str(tech_paragraph)
            TextLine = TechText.Add()
            TextLine.Str = edited_text
            TextLine.Numbering = 1
            if 'RGS 3.2.4.' in edited_text:
                tube_bending_flag = True
            if tube_bending_flag:
                if 'RGS 3.5.20' in edited_text:
                    TextLine.Str = 'ДОПУСКАЕТСЯ ИЗГОТАВЛИВАТЬ ИЗ ЧАСТЕЙ. ' + edited_text

        if tube_bending_flag is True:
            TextLine = TechText.Add()
            TextLine.Str = 'ИСПОЛЬЗУЕМЫЕ СОКРАЩЕНИЯ:'
            TextLine.Numbering = 1

            TextLine = TechText.Add()
            TextLine.Str = '    - ISO VIEW - ИЗОМЕТРИЧЕСКИЙ ВИД.'
            TextLine.Numbering = 0
            TextLine = TechText.Add()
            TextLine.Str = '    - NPS - НОМИНАЛЬНЫЙ РАЗМЕР ТРУБЫ.'
            TextLine.Numbering = 0
            TextLine = TechText.Add()
            TextLine.Str = '    - OD - НАРУЖНЫЙ ДИАМЕТР ТРУБЫ.'
            TextLine.Numbering = 0
            TextLine = TechText.Add()
            TextLine.Str = '    - WT - ТОЛЩИНА СТЕНКИ ТРУБЫ.'
            TextLine.Numbering = 0
            TextLine = TechText.Add()
            TextLine.Str = '    - L - ДЛИНА ТРУБЫ.'
            TextLine.Numbering = 0
        else:
            TextLine = TechText.Add()
            TextLine.Str = 'ИСПОЛЬЗУЕМЫЕ СОКРАЩЕНИЯ:'
            TextLine.Numbering = 1
            TextLine = TechText.Add()
            TextLine.Str = '    - ISO VIEW - ИЗОМЕТРИЧЕСКИЙ ВИД.'
            TextLine.Numbering = 0


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
            if i == 0:
                iLayoutSheet.LayoutStyleNumber = 31
            else:
                iLayoutSheet.LayoutStyleNumber = 32
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
        text_to_edit = str_to_edit.strip()
        #если целиком фраза есть в словаре
        if text_to_edit in self.DICTIONARY.keys():
            edited_text = self.DICTIONARY[text_to_edit]
            if str_to_edit.startswith(' '):
                return ' '+edited_text
            if str_to_edit.endswith(' '):
                return edited_text+' '
            return edited_text

        #поиск частичного совпадения слов словаря и полученного значения
        for key in self.DICTIONARY.keys():
            if key in text_to_edit:
                text_to_edit = text_to_edit.replace(key, self.DICTIONARY[key])
                
        #если фраза состоит из строк
        if '\n' in text_to_edit:
            edited_text = ''
            for row in text_to_edit.split('\n'):
                edited_text += self.edit_single_str(row)
                edited_text += '\n'
            return edited_text

        return self.edit_symbol_str(text_to_edit)

    def edit_symbol_str(self, str_to_edit: str):
        text_to_edit = str_to_edit.strip()
        if text_to_edit in self.DICTIONARY.keys():
            edited_text = self.DICTIONARY[text_to_edit]
            if str_to_edit.startswith(' '):
                return ' '+edited_text
            if str_to_edit.endswith(' '):
                return edited_text+' '

            return edited_text

        #поиск частичного совпадения слов словаря и полученного значения
        for key in self.DICTIONARY.keys():
            if key in text_to_edit:
                text_to_edit = text_to_edit.replace(key, self.DICTIONARY[key])

        edited_text = self.edit_single_str(text_to_edit)

        return edited_text

    def edit_single_str(self, str_to_edit:str):
        split_text = str_to_edit.split(' ')
        edited_text = ''
        for split_item in split_text:
            if split_item == 'PLACES':
                edited_text = self.get_correct_form(split_text[0], split_text[1])

            elif split_item == 'TUBE':
                edited_text += 'ТРУБА'

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

