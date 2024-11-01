# -*- coding: utf-8 -*-

import os, sys
import json
import pythoncom
from win32com.client import Dispatch, gencache
from tkinter.messagebox import showinfo
import threading

class TimeoutException(Exception):
    pass

def timeout_handler():
    raise TimeoutException('Прошло 10 секунд, выполнение прервано')

 
def get_path():
    '''
    Получить путь в текущую директорию
    '''
    return os.path.dirname(os.path.abspath(__file__))

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')

    return os.path.join(base_path, relative_path)


def read_json(path_to_file):
    '''
    Считывает JSON файл из переданного пути
    и возвращает полученную информацию
    '''
    path = get_path()+path_to_file
    #path = os.path.dirname(os.path.abspath(__file__)) #получение пути к данному модулю
    #techdemands_path = path+path_to_file #получение пути к файлу-источнику ТТ
    with open(path,'r', encoding='utf-8') as Json_file:
        templates = json.load(Json_file)
    return (templates)

class KompasAPI:
    '''
    Класс для подключения к КОМПАС-3D.
    При super() наследовании передает основные интефейсы:
    module - API компаса, app - экземпляр Application,
    const - константы Компаса
    '''
    def __init__(self):
        self.module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        self.app = self.module.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(self.module.IApplication.CLSID, 
                                                                                                     pythoncom.IID_IDispatch))
        self.const = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants
        self.const2D = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
        if self.app.Visible is False:
            try:
                timer = threading.Timer(10.0, timeout_handler)
                timer.start()
                try:
                    showinfo('Информация',
                            'Компас-3D запущен программой NerpaAI')
                    self.app.Visible = True

                finally:
                    timer.cancel()
            except TimeoutException as e:
                showinfo('Информация',
                         'Произошла ошибка при запуске Компас-3D. Попробуйте запустить его самостоятельно')
            except Exception:
                showinfo('Информация',
                         'Произошла ошибка при запуске Компас-3D. Попробуйте запустить его самостоятельно')
            
    
    def get_part_dispatch(self):
        '''
        Возвращает указатель на интерфейс iPart7
        для сборки
        '''
        iKompasDocument = self.app.ActiveDocument
        if iKompasDocument:
            iKompasDocument3D = self.module.IKompasDocument3D(iKompasDocument)
            return iKompasDocument3D.TopPart
        self.app.MessageBoxEx('Ошибка получения указателя iPart7',
                              'Ошибка', 64)
    
    def get_bodies_array(self):
        '''
        Возвращает list указателей тел из сборки.
        Внутри встроена проверка на включению тела в спецификацию
        и проверку зеркальности
        '''
        iPart7 = self.get_part_dispatch()
        bodies_array = self.module.IFeature7(iPart7).ResultBodies

        bodies_dispatches = []
        if bodies_array is not None:
            if isinstance(bodies_array, tuple):
                for body_dispatch in bodies_array:
                    body_object = KompasItem(body_dispatch)
                    if body_dispatch.CreateSpcObjects is True and body_object.is_patterned() is False:
                        bodies_dispatches.append(body_dispatch)
                return bodies_dispatches
            else:
                if bodies_array.CreateSpcObjects is True:
                    return [bodies_array]
                else:
                    return None
        else:
            self.app.MessageBoxEx('В активном документе нет тел',
                                  'Ошибка', 64)
            return
    
    def get_progress_bar(self):
        '''
        Возвращает интерфейс ProgressBar
        '''
        progress_bar = self.app.ProgressBarIndicator
        return progress_bar
    
class KompasItem:
    '''
    Класс с удобными методами для обработки тел и компонентов.
    Для создания объекта этого класса нужно передать Dispatch тела или компонента.
    '''
    def __init__(self, dispatch):
        self.dispatch = dispatch
        self.pattern_words = ['Массив', 'Зеркальное']
        self.kAPI = KompasAPI() 
        self.module = self.kAPI.module
        self.app = self.kAPI.app

    def get_prp_value(self,ID): 
        '''
        Возвращает значение свойства по переданному ID. Формат ID - float
        '''
        iPropertyMng = self.module.IPropertyMng(self.app)
        iPropertyKeeper = self.module.IPropertyKeeper(self.dispatch)
        return iPropertyKeeper.GetPropertyValue(iPropertyMng.GetProperty(self.app.ActiveDocument, ID),'',True, True)[1]

    def set_prp_value(self,ID, PrpValue):
        '''
        Устанавливает значение свойства по переданному ID. Формат ID - float
        '''
        iPropertyMng = self.module.IPropertyMng(self.app)
        iPropertyKeeper = self.module.IPropertyKeeper(self.dispatch)
        set_prp = iPropertyKeeper.SetPropertyValue(iPropertyMng.GetProperty(self.app.ActiveDocument, ID), PrpValue, True)
        return set_prp

    def is_patterned(self): 
        '''
        Метод проверки получен объект массивом или нет
        '''
        try:
            sub_features = self.module.IFeature7(self.dispatch).SubFeatures(0, True, True)
            for sub in sub_features:
                return any(word in sub.Name for word in self.pattern_words)
        except:
            return False

