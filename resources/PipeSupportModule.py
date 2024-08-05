# -*- coding: utf-8 -*-

from .NerpaUtility import KompasAPI, get_path

import logging, yaml, datetime, getpass


class PipeSupportDots(KompasAPI):
    def __init__(self):
        super().__init__()
        self.dir_path = get_path()
        self.username = getpass.getuser()
        self.current_time_date = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")

        #LOGGING BLOCK
        self.log_name = 'LOG {} {} {}.log'.format(self.username, 
                                               self.current_time_date,
                                               self.__class__.__name__)
        self.logging = logging
        logging.basicConfig(level=logging.INFO,
                            filename='logs/{}'.format(self.log_name),
                            filemode= 'w',
                            format="%(asctime)s %(levelname)s %(message)s")

        try:
            self.doc = self.app.ActiveDocument
            if self.doc:
                if self.doc.DocumentType != 4:
                    return
            self.iPart7 = self.get_part_dispatch()
            self.lines_coords = []
            self.dots_coords = []

            self.get_lines_coords()
            if self.define_support_dot():
                self.set_support_dots()
        except:
            self.logging.warning('Произошла ошибка формата. Дальнейшая работа не продолжилась')
            return
  
    
    def get_lines_coords(self):
        self.logging.info('Начало работы модуля')

        iAuxGeomContainer = self.module.IAuxiliaryGeomContainer(self.iPart7)
        LineSegments = iAuxGeomContainer.LineSegments3D
        for i in range(LineSegments.Count):
            LineSegment = LineSegments.LineSegment3D(i)
            iModelObject1 = self.module.IModelObject1(LineSegment)
            MathObject = iModelObject1.MathObject
            iMathCurve = self.module.IMathCurve3D(MathObject)
            if iMathCurve.GetGabarit()[0] is True:
                gabarits = iMathCurve.GetGabarit()
                round_gabarits = [True, 0,0,0,0,0,0]
                for i, coord in enumerate(gabarits):
                    if i == 0:
                        continue
                    else:
                        round_gabarits[i] = round(coord,2)

                self.lines_coords.append(round_gabarits)
                
        self.logging.info('Получены координаты линий. Всего обработано линий:{}'.format(len(self.lines_coords)))
        return True
        

    def define_support_dot(self):
        if self.lines_coords:
            for i, line in enumerate(self.lines_coords):
                first_point = tuple(line[1:4])
                second_point = tuple(line[4:])
                for n, other_line in enumerate(self.lines_coords):
                    if i == n:
                        continue
                    if first_point in (tuple(other_line[1:4]), 
                                       tuple(other_line[4:])) and first_point not in self.dots_coords:
                            self.dots_coords.append(first_point)

                    if second_point in (tuple(other_line[1:4]), 
                                       tuple(other_line[4:])) and second_point not in self.dots_coords:
                            self.dots_coords.append(second_point)
            self.logging.info('Получены координаты точек для суппорта. Всего найдено точек:{}'.format(len(self.dots_coords)))
            return True
        
        self.app.MessageBoxEx('Ошибка', 'Ошибка', 64)
        self.logging.info('Произошла ошибка в работе в функции define_support_dots')

    def set_support_dots(self):
        iModelContainer = self.module.IModelContainer(self.iPart7)
        Points3D = iModelContainer.Points3D
        for dot_coords in self.dots_coords:
            Point3D = Points3D.Add()
            Point3D.Symbol = 8.0
            Point3D.X = dot_coords[0]
            Point3D.Y = dot_coords[1]
            Point3D.Z = dot_coords[2]
            Point3D.Update()
            iModelObject = self.module.IModelObject(Point3D)
            ColorParam7 = self.module.IColorParam7(iModelObject)
            ColorParam7.Color = 1254577.0
            iModelObject.Update()
        self.logging.info('Точки расставлены')
                

