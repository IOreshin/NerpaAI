from NerpaUtility import KompasAPI


class AutoBendFinder(KompasAPI):
    def __init__(self):
        super().__init__()

    def get_unhis_curves(self):
        iPart7 = self.get_part_dispatch()
        if iPart7:
            iGeomCont = self.module.IAuxiliaryGeomContainer(iPart7)
            UnhisCurves = iGeomCont.UnhistoredCurves3D
            curves_placements = []
            for i in range(UnhisCurves.Count):
                unhis_curve = UnhisCurves(i)
                iModelObject1 = self.module.IModelObject1(unhis_curve)
                MathObject = iModelObject1.MathObject
                iMathCurve3D = self.module.IMathCurve3D(MathObject)
                curves_placements.append(iMathCurve3D.GetGabarit())
            return curves_placements
        return False
    
    def get_tube_route(self):
        doc = self.app.ActiveDocument
        iKompasDocument3D = self.module.IKompasDocument3D(doc)
        iSelectionManager = iKompasDocument3D.SelectionManager
        selected_objects = iSelectionManager.SelectedObjects
        if selected_objects is not None and len(selected_objects) == 4:
            self.all_curves = self.get_unhis_curves()
            try:
                self.start_point = selected_objects[0].GetPoint()[1:4]
                self.second_point = selected_objects[1].GetPoint()[1:4]
                self.pre_last_point = selected_objects[2].GetPoint()[1:4]
                self.last_point = selected_objects[3].GetPoint()[1:4]

            except Exception:
                print('SELECT ERROR')
                return
            
            self.tube_route = []   
            self.tube_route.append(self.start_point)
            self.tube_route.append(self.second_point)

            try:
                find_flag = True
                previous_dot = self.second_point
                counter = 0
                while find_flag is True and counter < 10002:
                    for curve in self.all_curves:
                        counter += 1
                        if curve[1:4] == previous_dot and curve[4:] not in self.tube_route: #начало 1:4 лежит в предыдущей точке
                            self.tube_route.append(curve[4:])
                            previous_dot = curve[4:]
                        elif curve[4:] == previous_dot and curve[1:4] not in self.tube_route:
                            self.tube_route.append(curve[1:4])
                            previous_dot = curve[1:4]
                        if previous_dot == self.pre_last_point:
                            self.last_point_flag = False
            except Exception:
                print(1)
                return

            if counter > 1000:
                print('CANNOT FIND ROUTE')
                return
            
            if self.tube_route[-1] != self.last_point:
                self.tube_route.pop()
                self.tube_route.append(self.last_point)

            return self.tube_route

test = AutoBendFinder()
test.get_tube_route()







