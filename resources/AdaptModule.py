# -*- coding: utf-8 -*-

from .NerpaUtility import KompasAPI, KompasItem
from .PropertyMngModule import PropertyManager
from .ConstantsRGSH import SOURCE_PROPERTIES_ID, B_SPEC, mNPS, mSCH, mRGS_PIPING, BOM_MTO_IDs

class AdaptParameters(KompasAPI):
    '''
    Класс получения данных и их адаптации по правилам РГШ-ИСО.
    Для создания объекта класса нужно передать Dispatch сборки или детали.
    '''
    def __init__(self, dispatch):
        super().__init__()
        self.dispatch = dispatch 
        self.state_arr = (None, '', '-')

    def get_property_values(self):
        '''
        Функция получения исходных данных свойств.
        Возвращает словарь со свойствами
        '''
        #создание словаря свойств
        property_values = dict()
        #создание объекта класса KompasItem для использования функции get_prp_value 
        body = KompasItem(self.dispatch)
        #цикл прохода по списку исходных свойств
        for prop_ID in SOURCE_PROPERTIES_ID:
            value = body.get_prp_value(prop_ID[0])
            property_values[prop_ID[1]] = value
        return property_values

    def lookup_value(self, value, lookup_table):
        '''
        Вспомогательная функция поиска соотвествия в list данных
        '''
        for item in lookup_table:
            if value == item[0]:
                return item[1]
        return '-'

    def get_prp_parameter(self,parameter, multiplier=None, rounding = None):
        '''
        Функция преобразования значения в зависимости от условий
        '''
        if parameter not in self.state_arr:
            return round(multiplier*parameter, rounding) if multiplier else parameter
        return '-'

    def get_description(self, nps, od, wt, tube_l, profile_l, thickness, description):
        '''
        Функция составления Наименования для объектов
        '''
        if od not in self.state_arr: #DESC FOR TUBE
            if od > 70:
                b_desc = 'PIPE NPS {} OD={} WT={} L={}'.format(nps,od,wt,round(1000*tube_l))
                m_desc = 'PIPE'
            else:
                b_desc = 'TUBE NPS {} OD={} WT={} L={}'.format(nps,od,wt,round(1000*tube_l))
                m_desc = 'TUBE'
            m_dim = 'ASME 36.10M'
            return b_desc, m_desc, m_dim
        elif thickness not in self.state_arr: #DESC FOR PLATE
            b_desc = 'PLATE {}'.format(round(1000*thickness))
            m_desc = b_desc
            m_dim = 'EN 10210'
            return b_desc, m_desc, m_dim
        else: #DESC FOR PROFILE
            b_desc = '{} L={}'.format(description, round(1000*profile_l))
            m_desc = '{}'.format(description)
            if description != 'RAIL DIN 3015': #DESC FOR PROFILE
                m_dim = 'EN 10210'
                return b_desc, m_desc, m_dim
            elif description == 'RAIL DIN 3015':
                m_dim = 'DIN 3015'
                return b_desc, m_desc, m_dim
        return '-','-', '-'

    def get_material(self, od, l_profile, mat, mat_name):
        '''
        функция для определения материала
        '''
        if od not in self.state_arr:
            return mat
        elif l_profile not in self.state_arr:
            if mat_name == "RAIL DIN 3015":
                return 'A4-70'
            else:
                return 'S355'
        else:
            return '-'

    def get_b_spec(self, od, l_prof, cv, profile_name, mRGS):
        '''
        Функция для определения bSPEC объекта
        '''
        if od not in self.state_arr: #TUBE
            return mRGS
        elif l_prof not in self.state_arr: #MK
            if profile_name == "RAIL DIN 3015":
                return "DIN EN ISO 3506-1"
            elif cv == 'CV1': #PLATE
                return B_SPEC[1]
            else:
                return B_SPEC[2]

    def get_mst(self, l_profile, l_tube):
        '''
        Функция определения CV объекта
        '''
        body = KompasItem(self.dispatch)
        mST = body.get_prp_value(245666335656.0)
        if mST in self.state_arr and l_profile not in self.state_arr:
            return 'CV1'
        elif l_tube not in self.state_arr:
            return '-'
        else:
            return mST
        
    def get_mRGS(self, m_tube, od):
        for item in mRGS_PIPING:
            if ['{}'.format(m_tube), od] == item[:2]:
                return item[2]
        return '-'
    
    def get_bom_mto_params(self):
        '''
        Основная функция по преобразованию исходных свойств в
        свойства по требованиям RGSH ISO
        '''
        try:
            property_values = self.get_property_values()

            m_od = self.get_prp_parameter(property_values['OD'], 1000, 2)
            m_wt = self.get_prp_parameter(property_values['WT'], 1000, 2)
            m_id = round(1000*(property_values['OD']-2*property_values['WT']),2) if property_values['OD'] not in self.state_arr else '-'
            m_nps = self.lookup_value(property_values['OD'], mNPS)
            m_sch = self.lookup_value(property_values['OD'], mSCH)
            m_rgs = self.get_mRGS(property_values['PIPE_MAT'], property_values['OD'])
            m_t = property_values['T'] if property_values['T'] not in self.state_arr else '-'
            m_width = round(property_values['WIDTH']*1000,5) if property_values['WIDTH'] not in self.state_arr else '-'

            if property_values['OD'] not in self.state_arr: #TUBE
                m_unit, m_value, sort_value = ('m', round(property_values['L_TUBE'],1), property_values['L_TUBE'])
            elif property_values['L_PROFILE'] not in self.state_arr:
                if property_values['T'] not in self.state_arr: #PLATE
                    m_unit, m_value, sort_value = ('m2', round((
                        property_values['L_PROFILE']*property_values['WIDTH']),2), 
                        (property_values['L_PROFILE']*property_values['WIDTH']))
                else: #PROFILE
                    m_unit, m_value, sort_value = ('m', round(property_values['L_PROFILE'],1), 
                                                property_values['L_PROFILE'])

            b_desc, m_desc, m_dim = self.get_description(m_nps, m_od, m_wt, 
                                                        property_values['L_TUBE'], property_values['L_PROFILE'],
                                                        m_t, property_values['PROFILE_NAME'])
            
            m_mat = self.get_material(property_values['OD'], property_values['L_PROFILE'],
                                    property_values['PIPE_MAT'],property_values['PROFILE_NAME'])
            
            b_mat = m_mat
            m_st = self.get_mst(property_values['L_PROFILE'], property_values['L_TUBE'])

            b_spec = self.get_b_spec(property_values['OD'], property_values['L_PROFILE'],m_st,
                                    property_values['PROFILE_NAME'], m_rgs)
            
            mass = round(property_values['MASS'], 1)
            b_type = 'DETAIL' if property_values['bTYPE'] in self.state_arr else property_values['bTYPE']

            BOM_MTO_VALUES_NAMES = (('mNPS', m_nps), ('mDESC', m_desc),('bDESC', b_desc), ('mDIM', m_dim), 
                                    ('mMAT', m_mat), ('bMAT', b_mat), ('mOD', m_od),('mID', m_id),
                                    ('mWT', m_wt), ('mSCH', m_sch), ('mRGS', m_rgs), ('mUNIT', m_unit),
                                    ('mVALUE', m_value), ('bSPEC', b_spec), ('mST', m_st),
                                    ('mT', str(m_t)), ('mWIDTH', str(m_width)), ('SORT_VALUE', sort_value),
                                    ('MASS', mass), ('bTYPE', b_type),)
            
            BOM_MTO_VALUES = {name[0]:None for name in BOM_MTO_VALUES_NAMES}
            for key in BOM_MTO_VALUES:
                for value_data in BOM_MTO_VALUES_NAMES:
                    if key == value_data[0]:
                        BOM_MTO_VALUES[key] = value_data[1]
        
            return BOM_MTO_VALUES
        except Exception:
            self.app.MessageBoxEx('Ошибка в свойствах тел. Проверьте правильность создания тел в документе',
                                  'Ошибка')
            return
            
class AdaptAssy(KompasAPI):
    '''
    Класс для обработки сборки
    '''
    def __init__(self):
        super().__init__()
        self.iPart7 = self.get_part_dispatch()
        self.adapt_current_assy()

    def get_start_index(self):
        '''
        Метод для получения стартового индекса для
        расстановки обозначений тел
        '''
        #получение SAFEARRAY компонентов
        parts = self.iPart7.PartsEx(self.const.ksUniqueParts)
        i = 1
        if parts is not None:
            top_marking = self.iPart7.Marking
            for part in parts:
                #проверка принадлежности детали текущей сборке
                part_marking = part.Marking[:-4]
                if part_marking == top_marking:
                    i += 1
        return i
    
    def get_marking_info(self):
        '''
        Функция получения list с свойствами уникальных тел,
        которые сортируются в marking_arr
        '''
        bodies_parameters = []
        bodies = self.get_bodies_array()
        if bodies is not None:
            for body_dispatch in bodies:
                adapt_body = AdaptParameters(body_dispatch)
                values = adapt_body.get_bom_mto_params()
                if values not in bodies_parameters:
                    bodies_parameters.append(values)

            marking_arr = sorted(bodies_parameters, key = lambda point:
                                 (point['mOD'], point['mT'], point['mDESC'], point['SORT_VALUE'],
                                  point['mWIDTH'], point['MASS']))
            
            return marking_arr

    def adapt_current_assy(self):
        '''
        Метод для адаптации входящих в сборку тел
        '''
        #проверка является ли активный документ сборкой
        doc = self.app.ActiveDocument
        if doc.DocumentType != 5:
            self.app.MessageBoxEx('Активный файл не является сборкой',
                                  'Ошибка формата', 
                                  64)
            return
        
        #проверка и добавление свойств RGSH
        property_manager = PropertyManager()
        property_manager.check_add_properties()

        MARKING_INFO = self.get_marking_info() #лист с инфо для маркировки
        START_INDEX = self.get_start_index() #стартовый индекс для маркировки
        if MARKING_INFO:
            progress_bar = self.get_progress_bar()
            progress_bar.Start(0.0, len(MARKING_INFO),'', True)
            bodies = self.get_bodies_array()
            for i,body_dispatch in enumerate(bodies):
                body_parameters = AdaptParameters(body_dispatch)
                bom_mto_params = body_parameters.get_bom_mto_params()
                if bom_mto_params['mVALUE'] != '-':
                    body_object = KompasItem(body_dispatch)

                    for id in BOM_MTO_IDs:
                        if bom_mto_params[id[1]]:
                            body_object.set_prp_value(id[0], bom_mto_params[id[1]])

                    drw_format = body_object.get_prp_value(1.0)
                    if drw_format in ['БЧ', '', None, '-']:
                        body_object.set_prp_value(1.0, 'N/D')
                    
                    if bom_mto_params['mMAT'] != '-':
                        body_object.set_prp_value(9.0, bom_mto_params['mMAT'])
                    
                    marking_index = MARKING_INFO.index(bom_mto_params)+START_INDEX
                    marking_body = str(self.iPart7.Marking)+'-'+'{:03}'.format(marking_index)
                    body_object.set_prp_value(4.0, marking_body)

                progress_bar.SetProgress(i,'', True)
            progress_bar.Stop('', True)
            self.app.MessageBoxEx('Объекты адаптированы. Проверьте корректность выполнения операции.',
                                  'Успех!', 64)
            
class AdaltDetail(KompasAPI):
    '''
    Класс для обработки детали
    '''
    def __init__(self):
        super().__init__()
        self.adapt_current_detail()

    def adapt_current_detail(self):
        '''
        Метод для обработки детали
        '''
        #проверка формата документа
        doc = self.app.ActiveDocument
        if doc.DocumentType != 4:
            self.app.MessageBoxEx('Активный документ не является деталью',
                                  'Ошибка формата', 64)
            return
        
        iPart7 = self.get_part_dispatch()
        bodies = self.module.IFeature7(iPart7).ResultBodies
        if bodies is not None:
            true_bodies = []
            if isinstance(bodies, tuple):
                for body in bodies:
                    body_object = KompasItem(body)
                    if body_object.get_prp_value(8.0) != 0:
                        true_bodies.append(body)
            else:
                body_object = KompasItem(bodies)
                if body_object.get_prp_value(8.0) != 0:
                    true_bodies.append(bodies)
        else:
            self.app.MessageBoxEx('В детали нет тел',
                                  'Ошибка', 64)
            return
        
        if len(true_bodies) != 1:
            self.app.MessageBoxEx('Ошибка', 
                     'Деталь содержит 2 или более тел', 64)
            return
        
        #проверка и добавление свойств RGSH в деталь
        property_manager = PropertyManager()
        property_manager.check_add_properties()

        adapt_body = AdaptParameters(true_bodies[0])
        bom_mto_params = adapt_body.get_bom_mto_params()
        iPart7 = self.get_part_dispatch()
        detail_object = KompasItem(iPart7)
        for id in BOM_MTO_IDs:
            if id[1] == 'bDESC':
                continue
            if bom_mto_params[id[1]]:
                detail_object.set_prp_value(id[0], bom_mto_params[id[1]])
            iPart7.Update()

