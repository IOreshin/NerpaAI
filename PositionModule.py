# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Назначение модуля: реализованы функции для установки позиций объектам сборки.
# Практика показала, что спецификация неадекватно работает со сборкой
# металлоконструкций и трубопроводов, в результате чего при расстановке позиций
# из активной сборки вырезаются любые спецификации и позиции устанавливаются
# программно.
#-------------------------------------------------------------------------------

from NerpaUtility import KompasAPI, format_error, KompasItem
from ConstantsRGSH import POSITION_ID, bTYPE_NAMES
from tkinter.messagebox import showinfo


class SetPositions(KompasAPI):
    '''
    Класс для устновки позиций для тел и компонентов
    '''
    def __init__(self):
        super().__init__() #получение интерфейсов KompasAPI
        self.iPart7 = self.get_part_dispatch() #получение интерфейса iPart7

    def remove_spec_desc(self):
        '''
        Метод удаления описаний спецификаций в сборке
        '''
        iKompasDocument = self.app.ActiveDocument #интерфейс активного окно
        if iKompasDocument.DocumentType != 5: #если не сборка
            self.app.MessageBoxEx('Активный документ не является сборкой',
                                  'Ошибка', 64)
            return
        SpecDescs = iKompasDocument.SpecificationDescriptions #коллекция спецификаций
        spec = SpecDescs.Active #попытка получить спецификацию
        if spec:
            while spec is not None:
                spec.Delete()
                spec = SpecDescs.Active
        
        return True

    def get_item_params(self, dispatch):
        '''
        Функция сбора данных для сортировки
        '''
        item_params = [] #массив данных для сортировки
        for ID in POSITION_ID: 
            item = KompasItem(dispatch) #создание объекта для функций
            prop_value = item.get_prp_value(ID)
            if prop_value is None:
                prop_value = ''
            item_params.append(prop_value)

        try:
            Index = POSITION_ID.index(item_params[0])
            item_params[0] = Index
        except Exception:
            item_params[0] = len(bTYPE_NAMES)+1

        return tuple(item_params)
    
    def set_item_pos(self, parts_list:list, dispatch): 
        '''
        функция установки значения позиции объекту
        '''
        item_params = self.get_item_params(dispatch)
        pos = parts_list.index(item_params)
        item = KompasItem(dispatch)
        item.set_prp_value(15.0, (pos+1))

    def get_parts_list(self):
        '''
        Функция сбора объектов для позиций
        '''
        parts_list = []
        parts = self.iPart7.PartsEx(self.const.ksUniqueParts)
        if parts is not None:
            for part in parts:
                if part.IsLayoutGeometry is False and part.CreateSpcObjects is True:
                    part_params = self.get_item_params(part)
                    parts_list.append(part_params)

        bodies = self.module.IFeature7(self.iPart7).ResultBodies
        if bodies is not None:
            if isinstance(bodies, tuple):
                for body in bodies:
                    if body.CreateSpcObjects:
                        body_params = self.get_item_params(body)
                        parts_list.append(body_params)
            else:
                if bodies.CreateSpcObjects:
                    body_params = self.get_item_params(body)
                    parts_list.append(body_params)
        return parts_list

    def sort_set_positions(self): 
        '''
        Функция сортировки-установки значения позиции
        '''
        parts_list = list(set(self.get_parts_list()))
        #сортировка производится в следующем порядке: bTYPE, PN, DESC
        parts_list = sorted(parts_list, key= lambda x:(x[0], x[1], x[2]))

        progress_bar = self.get_progress_bar()
        progress_bar.Start(0.0, len(parts_list), '', True)

        #установка позиций компонентам
        parts = self.iPart7.PartsEx(self.const.ksAllParts)
        if parts is not None:
            progress_bar.start(0.0, len(parts),'',True)
            for i, part in enumerate(parts):
                if part.IsLayoutGeometry is not False and part.CreateSpcObjects:
                    self.set_item_pos(parts_list, part)
                    progress_bar.SetProgress(i,'',True)
            progress_bar.Stop('', True)

        #установка позиций телам
        bodies = self.module.IFeature7(self.iPart7).ResultBodies
        if bodies is not None:
            if isinstance(bodies,tuple):
                progress_bar.Start(0.0,len(bodies),'', True)
                for i, body in enumerate(bodies):
                    if body.CreateSpcObjects:
                        self.set_item_pos(parts_list, body)
                        progress_bar.SetProgress(i,'',True)
                progress_bar.Stop('', True)
            else:
                if bodies.CreateSpcObjects:
                    self.set_item_pos(parts_list, bodies)

    def set_positions(self):
        if self.remove_spec_desc():
            self.sort_set_positions()
            self.app.MessageBoxEx('Позиции установлены. Проверьте корректность выполнения операции',
                                    'Успех', 64)
