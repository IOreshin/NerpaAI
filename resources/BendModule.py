# -*- coding: utf-8 -*-

from .NerpaUtility import KompasAPI, KompasItem, read_json
from .WindowModule import Window
import math
import os
import tkinter as tk
from .AutoBendModule import AutoBendFinder
from tkinter import ttk


def select_error():
    KompasAPI().app.MessageBoxEx('Среди выделенных объектов есть прочие объекты либо нарушен порядок выбора. Проверьте выделение и повторите команду',
                                 'Ошибка', 64)
    
#класс окна-интерфейса
class BTWindow(Window):
    def __init__(self):
        super().__init__()
        self.method_combobox = None
        self.product_combobox = None
        self.products_names = ('KM',
                              'SDU',
                              'CM')
        self.methods_names = ('По линиям',
                                  'По точкам',
                                  'Авто')
        self.get_bt_window()

    def get_bt_window(self):
        root = tk.Tk()
        root.title(self.window_name+' : BEND TABLE')
        root.iconbitmap(self.pic_path)
        root.resizable(False, False)
        root.attributes("-topmost", True)
        
        frame1 = ttk.LabelFrame(root, borderwidth = 5, relief = 'solid', text = 'Режим работы')
        frame1.grid(row = 0, column = 0, pady = 5, padx = 5, sticky = 'nsew')

        frame2 = ttk.LabelFrame(root, borderwidth = 5, relief = 'solid', text = '')
        frame2.grid(row = 1, column = 0, pady = 5, padx = 5, sticky = 'nsew')

        Label1 = tk.Label(frame1, text = "Изделие", state=["normal"])
        Label1.grid(row = 0, column = 0, sticky = 'nsew', pady = 2, padx = 5)
        Label2 = tk.Label(frame1, text = "Способ построения", state=["normal"])
        Label2.grid(row = 0, column = 1, sticky = 'nsew', pady = 2, padx = 5)

        
        self.product_combobox = ttk.Combobox(frame1, 
                                            values= self.products_names)
        self.product_combobox.current(0)
        self.product_combobox.grid(row = 1, column = 0,
                                   padx=5,pady=5)

        self.method_combobox = ttk.Combobox(frame1, 
                                 values= self.methods_names)
        self.method_combobox.current(1)
        self.method_combobox.grid(row = 1, column=1, padx=5,pady=5)

        def form_bend_info():
            bend_table_calculator = BendTableCalculator(self)
            bend_table_calculator.get_coordinate_info()

        def write_bend_info():
            bend_table_writer = BendTableWriter()
            bend_table_writer.write_bend_table_doc()

        def help_page():
            help_page = BTHelpWindow()
            help_page.get_help_window()

        self.create_button(ttk, frame2, 'Сформировать', form_bend_info, 14, 'normal', 0, 0)
        self.create_button(ttk, frame2, 'Добавить', write_bend_info, 14, 'normal', 0, 1)
        self.create_button(ttk, frame2, 'Справка', help_page, 14, 'normal',0,2)

        root.update_idletasks()
        w,h = self.get_center_window(root)
        root.geometry('+{}+{}'.format(w,h))
        root.mainloop()

class BendTableCalculator(KompasAPI):
    def __init__(self, window_instance):
        super().__init__()
        self.window_instance = window_instance

    def dot_operation(self, dot_dispatch, coordinate_list):
        '''
        Функция обработки точки. Для работы нужен диспатч iVertex и лист с координатами
        внутри есть проверка того точка это или нет и округление до одного знака
        '''
        try:
            point_coordinate = list(dot_dispatch.GetPoint()) #[BOOL,X1,Y1,Z1]
            if point_coordinate[0] is not True:
                select_error()
                return
            for i in range(3):
                point_coordinate[i+1] = round(point_coordinate[i+1],1)
            coordinate_list.append(point_coordinate[1:])
            return True
        
        except Exception:
            return False

    def curve_operation(self, current_curve, next_curve, last_coordinate, coordinate_list):
        '''
        Функция обработки кривой. Для работы нужен диспатч iEdge, диспатч следующей кривой iEdge,
    последнюю добавленную координату и лист с координатами
    Внутри реализованы проверка кривая это или нет, округление до 1 знака,
    случай, когда начало и конец отрезка не лежит в последней координате(тройник например)
        '''
        try:
            iMathCurve3D = current_curve.MathCurve 
            placement_matrix = list(iMathCurve3D.GetGabarit()) #[BOOL, x1,y1,z1,x2,y2,z2]
            if placement_matrix[0] is not True: #проверка кривая это или нет
                select_error()
                return
            for i in range(6): 
                placement_matrix[i+1] = round(placement_matrix[i+1], 1)
            
            if placement_matrix[1:4] == last_coordinate:
                coordinate_list.append(placement_matrix[4:])
                return
            
            elif placement_matrix[4:] == last_coordinate:
                coordinate_list.append(placement_matrix[1:4])
                return
            
            else:#случай, когда ТОЧКА начала или конца трубы не лежит в начале или конце ОТРЕЗКА
                iMathCurve3D_next_curve = next_curve.MathCurve
                next_placement_matrix = list(iMathCurve3D_next_curve.GetGabarit())
                if next_placement_matrix[0] is not True:
                    select_error()
                    return
                
                for i in range(6):
                    next_placement_matrix[i+1] = round(next_placement_matrix[i+1],1)

                next_placement = [next_placement_matrix[1:4], next_placement_matrix[4:]]

                if placement_matrix[1:4] in next_placement:
                    coordinate_list.append(placement_matrix[1:4])
                    return
                
                elif placement_matrix[4:] in next_placement:
                    coordinate_list.append(placement_matrix[4:])
                    return
                
                
        except Exception:
            select_error()
            return
            
    def get_coordinate_info(self):
        try:
            self.doc = self.app.ActiveDocument
            if self.doc.DocumentType != 5:
                self.app.MessageBoxEx('Активный документ не является сборкой',
                                    'Ошибка', 64)
                return
        except Exception:
            self.app.MessageBoxEx('Активный документ не является сборкой',
                                    'Ошибка', 64)
            return
        
        iKompasDocument3D = self.module.IKompasDocument3D(self.doc)
        iSelectionManager = iKompasDocument3D.SelectionManager
        selected_objects = iSelectionManager.SelectedObjects
        if selected_objects is None or isinstance(selected_objects, tuple) is False:
            self.app.MessageBoxEx('Должно быть выделено более 2 точек',
                                  'Ошибка', 64)
            return
        
        source_dots_coords = [] #массив изначальных координат
        if self.window_instance.method_combobox.get() == self.window_instance.methods_names[0]: #СПОСОБ ПО ЛИНИЯМ
            for i, object in enumerate(selected_objects):
                try:
                    if i == 0: #подразумевается,
                        #что первый и последний выделенный объект-точки
                        self.dot_operation(object, source_dots_coords)
                    elif i == (len(selected_objects)-1):
                        self.dot_operation(object, source_dots_coords)
                    elif i == (len(selected_objects)-2): #не учитываем последний отрезок
                        continue
                    else:#подразумевается операция над кривой
                        self.curve_operation(object, selected_objects[i+1], source_dots_coords[i-1],source_dots_coords)
                except Exception:
                    select_error()
                    return
                
        elif self.window_instance.method_combobox.get() == self.window_instance.methods_names[1]: #СПОСОБ ПО ТОЧКАМ
            try:
                for i, object in enumerate(selected_objects):
                    state = self.dot_operation(object, source_dots_coords)
                    if state is False:
                        select_error()
                        return
            except Exception:
                select_error()
                return
        
        elif self.window_instance.method_combobox.get() == self.window_instance.methods_names[2]: #АВТО СПОСОБ
            try:
                auto_finder = AutoBendFinder()
                source_dots_coords = auto_finder.get_tube_route()
            except Exception:
                select_error()
                return
            
        first_coordinates = source_dots_coords[0] #запоминаем первую точку,
        #нужно для обнуления
        zero_dots_coords = [] #массив координат в нулевой системе координат

        for i, point_coord in enumerate(source_dots_coords):
            if i == 0:
                zero_dots_coords.append([0,0,0])
            else:
                if self.window_instance.product_combobox.get() == self.window_instance.products_names[2]: #если CM
                    current_point_coord = []
                    for n in range(3):
                        coordinate = round(point_coord[n]-first_coordinates[n])
                        current_point_coord.append(coordinate)
                    zero_dots_coords.append(current_point_coord)

                elif self.window_instance.product_combobox.get() == self.window_instance.products_names[0]: #если выбран КМ и требуется корректировка
                    current_point_coord = [0,0,0]
                    for n in range(3):
                        coordinate = round((point_coord[n]-first_coordinates[n]))
                        if n == 0:
                            current_point_coord[1] = -coordinate
                        elif n == 1:
                            current_point_coord[2] = -coordinate
                        else:
                            current_point_coord[0] = coordinate
                        
                    zero_dots_coords.append(current_point_coord)

                else: #если выбран SDU
                    current_point_coord = [0,0,0]
                    for n in range(3):
                        coordinate = round((point_coord[n]-first_coordinates[n]))
                        if n == 0:
                            current_point_coord[0] = -coordinate
                        elif n == 1:
                            current_point_coord[1] = -coordinate
                        else:
                            current_point_coord[2] = coordinate
                        
                    zero_dots_coords.append(current_point_coord)

        relative_point_coord = [] #массив смещения координат
        for i, point_coordinate in enumerate(zero_dots_coords):
            if i == 0:
                relative_point_coord.append([0,0,0])
            else:
                point_movement = []
                for n, coordinate in enumerate(point_coordinate):
                    if zero_dots_coords[i][n] == zero_dots_coords[i-1][n]:
                        point_movement.append(0)
                    else:
                        point_movement.append(zero_dots_coords[i][n]-zero_dots_coords[i-1][n])
                relative_point_coord.append(point_movement)

        bend_radius = 75

        def calculate_ybc(points, bend_radius):
            def distance(p1, p2):
                return math.sqrt((p2[0] - p1[0])**2 + (p2[1] - p1[1])**2 + (p2[2] - p1[2])**2)
            
            def angle_between_vectors(v1, v2):
                dot_product = sum(v1[i] * v2[i] for i in range(3))
                mag_v1 = math.sqrt(sum(v1[i] ** 2 for i in range(3)))
                mag_v2 = math.sqrt(sum(v2[i] ** 2 for i in range(3)))
                if mag_v1 == 0 or mag_v2 == 0:
                    return 0
                cos_theta = max(min(dot_product / (mag_v1 * mag_v2), 1), -1)
                return math.acos(cos_theta) * (180.0 / math.pi)
            
            def cross_product(v1, v2):
                return [
                    v1[1] * v2[2] - v1[2] * v2[1],
                    v1[2] * v2[0] - v1[0] * v2[2],
                    v1[0] * v2[1] - v1[1] * v2[0]
                ]
            
            def normalize(v):
                mag = math.sqrt(sum(coord ** 2 for coord in v))
                return [coord / mag for coord in v] if mag != 0 else [0, 0, 0]
            
            def angle_between_planes(n1, n2):
                dot_product = sum(n1[i] * n2[i] for i in range(3))
                mag_n1 = math.sqrt(sum(n1[i] ** 2 for i in range(3)))
                mag_n2 = math.sqrt(sum(n2[i] ** 2 for i in range(3)))
                if mag_n1 == 0 or mag_n2 == 0:
                    return 0
                cos_theta = max(min(dot_product / (mag_n1 * mag_n2), 1), -1)
                return math.acos(cos_theta) * (180.0 / math.pi)

            ybc_data = []
            previous_normal = None
            
            for i in range(1, len(points) - 1):
                total_distance = distance(points[i - 1], points[i])
                v1 = [points[i][j] - points[i - 1][j] for j in range(3)]
                v2 = [points[i + 1][j] - points[i][j] for j in range(3)]
                C = angle_between_vectors(v1, v2)
                
                current_normal = normalize(cross_product(v1, v2))
                
                if previous_normal is None:
                    B = 0
                else:
                    B = angle_between_planes(previous_normal, current_normal)
                    # Определить направление угла поворота
                    cross_prod = cross_product(previous_normal, current_normal)
                    if cross_prod[2] < 0:
                        B = -B

                previous_normal = current_normal
                
                # Calculate arc length of the bend (R * θ)
                arc_length = bend_radius * math.tan(math.radians(C)/2) ##это не длина дуги а её проекция
                
                # Calculate the offset (effective straight distance between bends excluding arc length)
                if i == 1:
                    offset = total_distance - bend_radius*math.tan(math.radians(C)/2)

                elif i > 1 and i < len(points)-1:
                    arc_length_next = bend_radius*math.tan(math.radians(ybc_data[i-2][2])/2)
                    offset = total_distance - arc_length-arc_length_next
                
                ybc_data.append((offset, B, C))
                
            # Adding the last segment's straight distance
            Y_last = distance(points[-2], points[-1]) - bend_radius*math.tan(math.radians(ybc_data[i-1][2])/2)
            ybc_data.append((Y_last, 0, 0))
            
            return ybc_data
        
        ybc_result = calculate_ybc(zero_dots_coords, bend_radius)

        #TODO : надо это прибрать как-то
        for i, value in enumerate(ybc_result):
            if i == 0:
                relative_point_coord[i].append(0)
                relative_point_coord[i].append(0)
                relative_point_coord[i].append(0)
                relative_point_coord[i].append(0)
                relative_point_coord[i+1].append(bend_radius)
                relative_point_coord[i+1].append(round(value[2]))
                relative_point_coord[i+1].append(round(value[1]))
                relative_point_coord[i+1].append(round(value[0]))
            else:
                relative_point_coord[i+1].append(bend_radius)
                relative_point_coord[i+1].append(round(value[2]))
                relative_point_coord[i+1].append(round(value[1]))
                relative_point_coord[i+1].append(round(value[0]))

        global BendTempFilePath
        BendTempFilePath = ('{}\\BENDTEMP.txt').format(self.doc.Path)

        with open(BendTempFilePath, 'w') as temp_file:
            for coordinates in relative_point_coord:
                for coordinate in coordinates:
                    temp_file.write(str(coordinate)+',')
                temp_file.write('\n')

        self.app.MessageBoxEx('BEND TABLE сформирована. Перейдите в чертеж выделенной трубы и нажмите кнопку <Добавить>',
                              'Успех', 64)

class BendTableWriter(KompasAPI):
    def __init__(self):
        super().__init__()
    
    def write_bend_table_doc(self):
        try:
            self.doc = self.app.ActiveDocument
            if self.doc.DocumentType != 1:
                self.app.MessageBoxEx('Активный документ не является чертежом',
                                    'Ошибка', 64)
                return
        except:
            self.app.MessageBoxEx('Активный документ не является чертежом',
                                    'Ошибка', 64)
            return

        
        #проверка наличия видов модели в чертеже
        iKompasDocument2D = self.module.IKompasDocument2D(self.doc)
        ViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager
        iViews = ViewsAndLayersManager.Views
        if iViews.Count == 1:
            self.app.MessageBoxEx('В чертеже нет видов модели',
                                  'Ошибка', 64)
            return
        
        #проверка наличия вида BEND TABLE
        for i in range(iViews.Count):
            iView = iViews.View(i)
            if iView.Name == 'BEND TABLE':
                self.app.MessageBoxEx('В чертеже уже есть вид с BEND TABLE',
                                  'Ошибка', 64)
                return
        
        #считывания сформированных координат
        try: 
            with open(BendTempFilePath, 'r') as TempFile: #чтение JSON файла с координатами
                BendInfo = TempFile.readlines() #получение массива данных-координат
        except Exception:
            self.app.MessageBoxEx('Отсутствует сформированная BEND TABLE. Перейдите в сборку трубопровода и следуя инструкции выделите объекты трубы',
                                  'Ошибка', 64)
            return
        
        #получение из названия трубы полной длины.
        #конструкция try-except сделана на случай, когда в названии может не быть L= и тд.
        try:
            drw_object = KompasItem(self.doc)
            desc_tube_drw = drw_object.get_prp_value(5.0)
            total_length = desc_tube_drw.split('L=',1)[1].lstrip()
        except Exception:
            self.app.MessageBoxEx('В Наименовании чертежа некорректная запись трубы. Перед снятием чертежа трубы проведите адаптацию сборки',
                                  'Ошибка', 64)
            
        #получение формата листа и координат размещения вида BEND TABLE
        iLayoutSheets = self.doc.LayoutSheets
        iLayoutSheet = iLayoutSheets.Item(0)
        iSheetFormat = iLayoutSheet.Format
        Format = iSheetFormat.Format
        iView = iViews.Add(1)
        iDrawingObject = self.module.IDrawingObject(iView)
        iView.Current = True
        iView.Name = "BEND TABLE"
        formats = [['A0', 1184, 836],
                    ['A1', 836, 589],
                    ['A2', 589, 415],
                    ['A3', 415, 292],
                    ['A4', 205, 292],]
        iView.X = formats[Format][1]
        iView.Y = formats[Format][2]
        iView.Update()
        iDrawingObject.Update()

        #создание таблицы в виде BEND TABLE
        iSymbols2DContainer = self.module.ISymbols2DContainer(iView)
        DrawingTables = iSymbols2DContainer.DrawingTables
        TableHeight = 7
        TableWidth = 25
        TitleNames = ('POINT №', 'X', 'Y', 'Z', 'BEND@/RADIUS', 'BEND@/ANGLE', 'TWIST@/ANGLE', 'OFFSET')
        BendTable = DrawingTables.Add(1,len(TitleNames),TableHeight,TableWidth,0)
        iTable = self.module.ITable(BendTable)
        iTable.AddRow(0, True)


        #передача в таблицу названия столбцов и данных
        for i, Name in enumerate(TitleNames):
            iTableCell = iTable.Cell(1, i)
            iText = self.module.IText(iTableCell.Text)
            iText.Str = Name

        for i, Coordinates in enumerate(BendInfo):
            iTable.AddRow(i+1, True)
            iTableCell = iTable.Cell(i+2, 0)
            iText = self.module.IText(iTableCell.Text)
            iText.Str = i+1
            Values = Coordinates.split(',')
            Values.pop()
            for n, Value in enumerate(Values):
                iTableCell = iTable.Cell(i+2, n+1)
                iText = self.module.IText(iTableCell.Text)
                iText.Str = Value
        
        #создание главного заголовка
        FirstRow = iTable.Range(0,0,0,len(TitleNames))
        FirstRow.CombineCells()
        FirstRowCell = iTable.Cell(0,0)
        iText = self.module.IText(FirstRowCell.Text)
        iText.Str = 'BEND LOCATION INFORMATION'

        #создание последней строки с указание длины трубы "TOTAL LENGTH XXXX"
        iTable.AddRow(iTable.RowsCount-1, True)
        LastRowA = iTable.Range(iTable.RowsCount-1,0,iTable.RowsCount-1,len(TitleNames)-3)
        LastRowA.CombineCells()
        LastRowB = iTable.Range(iTable.RowsCount-1,1,iTable.RowsCount-1,len(TitleNames))
        LastRowB.CombineCells()
        LastRowCellA = iTable.Cell(iTable.RowsCount-1, 0)
        iText = self.module.IText(LastRowCellA.Text)
        iText.Str = 'TOTAL LENGTH'
        LastRowCellB = iTable.Cell(iTable.RowsCount-1, 6)
        iText = self.module.IText(LastRowCellB.Text)
        iText.Str = total_length

        #добавим стиля таблице (границы ячеек)
        WholeTable = iTable.Range(0, 0, iTable.RowsCount-1, len(TitleNames)-1) #Выделяем все ячейки
        Bound = WholeTable.CellsBoundaries
        iBound = self.module.ICellBoundaries(Bound)
        iBound.SetLineStyle(self.const2D.ksCBAllBorders, 1) #пришлось вызвать контанты2D для "ksCBAllBorders"
        subTitle = iTable.Cell(1, 0) #Теперь раздвинем шапку на две строчки (хитрим и работаем только через "одну" ячейку)
        iCFormat = self.module.ICellFormat(subTitle)
        iCFormat.Height = 14.0

        #размещение таблицы в виде
        BendTable.X = -TableWidth*len(TitleNames)
        BendTable.Y = 0
        BendTable.Update()
        
        #удаление временного файла и объявление успешности работы
        os.remove(BendTempFilePath)
        self.app.MessageBoxEx('BEND TABLE добавлен в чертеж. Проверьте корректность выполнения операции',
                              'Успех', 64)
        return

class BTHelpWindow(Window):
    def __init__(self):
        super().__init__()

    def get_help_window(self):
        help_root = tk.Tk()
        help_root.title(self.window_name+' : Справка')
        help_root.iconbitmap(self.pic_path)
        help_root.resizable(False, False)
        help_root.attributes("-topmost", True)
        
        help_frame = ttk.LabelFrame(help_root, borderwidth = 5, relief = 'solid', text = 'Справка')
        help_frame.grid(row = 0, column = 0, pady = 5, padx = 5, sticky = 'nsew')
        
        HelpPages = read_json('resources/lib/HELPPAGES.json' )
        HelpText = ''
        for item, text in HelpPages.items():
            if item == 'BEND TABLE HELP':
                for stroka in text:
                    HelpText = HelpText+stroka+'\n'

        help_label = tk.Label(help_frame, text = HelpText, state=["normal"])
        help_label.grid(row = 0, column = 0, sticky = 'e', pady = 2, padx = 5)

