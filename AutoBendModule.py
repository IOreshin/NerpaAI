# -*- coding: utf-8 -*-

from NerpaUtility import KompasAPI
from collections import defaultdict, deque

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
    
    def is_point_on_curve(self, A, B, P, tol=1e-9):
        x1, y1, z1 = A
        x2, y2, z2 = B
        x, y, z = P

        def close(a, b, tol):
            return abs(a - b) <= tol

        # Векторные операции
        AB = (x2 - x1, y2 - y1, z2 - z1)
        AP = (x - x1, y - y1, z - z1)
        
        # Векторное произведение
        cross_product = (
            AB[1] * AP[2] - AB[2] * AP[1],  # x компонента векторного произведения
            AB[2] * AP[0] - AB[0] * AP[2],  # y компонента векторного произведения
            AB[0] * AP[1] - AB[1] * AP[0]   # z компонента векторного произведения
        )
        
        if not (close(cross_product[0], 0, tol) and 
                close(cross_product[1], 0, tol) and
                close(cross_product[2], 0, tol)):
            return False
        
        # Проверка, находится ли точка P в пределах отрезка
        if min(x1, x2) - tol <= x <= max(x1, x2) + tol and \
        min(y1, y2) - tol <= y <= max(y1, y2) + tol and \
        min(z1, z2) - tol <= z <= max(z1, z2) + tol:
            return True

        return False

    def round_coord(self, values):
        for i, value in enumerate(values):
            values[i] = round(value, 1)

        return values

    def get_tube_route(self):
        doc = self.app.ActiveDocument
        iKompasDocument3D = self.module.IKompasDocument3D(doc)
        iSelectionManager = iKompasDocument3D.SelectionManager
        selected_objects = iSelectionManager.SelectedObjects
        if selected_objects is not None and len(selected_objects) == 2:
            self.all_curves = self.get_unhis_curves()
            try:
                self.start_point = tuple(self.round_coord(
                    list(selected_objects[0].GetPoint()[1:4])))
                self.last_point = tuple(self.round_coord(
                    list(selected_objects[1].GetPoint()[1:4])))
                
            except Exception:
                print('SELECT ERROR')
                return
            
            graph = self.build_graph(self.all_curves,
                                     self.start_point,
                                     self.last_point)
            
            path = self.bfs_path(graph, self.start_point,
                                 self.last_point)
            for i,coords in enumerate(path):
                path[i] = list(coords)

            return path
            
    def build_graph(self, curves, start, end):
        """
        Строит граф из кривых.

        :param curves: Список кривых, где каждая кривая задается кортежем (is_connected, x1, y1, z1, x2, y2, z2).
        :param start: Начальная точка (x, y, z).
        :param end: Конечная точка (x, y, z).
        :return: Граф в виде словаря, где ключи — точки, а значения — список смежных точек.
        """
        graph = defaultdict(list)
        points = set()
        
        # Добавляем начальную и конечную точки
        points.add(start)
        points.add(end)

        # Добавляем точки и рёбра из кривых
        for curve in curves:
            A = (curve[1], curve[2], curve[3])
            B = (curve[4], curve[5], curve[6])
            points.add(A)
            points.add(B)
            graph[A].append(B)
            graph[B].append(A)
        
        # Проверяем, лежат ли начальная и конечная точки на каких-либо отрезках
        for curve in curves:
            A = (curve[1], curve[2], curve[3])
            B = (curve[4], curve[5], curve[6])
            if self.is_point_on_curve(A, B, start):
                points.add(start)
                if A not in graph[start]:
                    graph[start].append(A)
                if B not in graph[start]:
                    graph[start].append(B)
                if A not in graph:
                    graph[A] = []
                if start not in graph[A]:
                    graph[A].append(start)
                if B not in graph:
                    graph[B] = []
                if start not in graph[B]:
                    graph[B].append(start)
            if self.is_point_on_curve(A, B, end):
                points.add(end)
                if end not in graph:
                    graph[end] = []
                if A not in graph[end]:
                    graph[end].append(A)
                if B not in graph[end]:
                    graph[end].append(B)
                if A not in graph:
                    graph[A] = []
                if end not in graph[A]:
                    graph[A].append(end)
                if B not in graph:
                    graph[B] = []
                if end not in graph[B]:
                    graph[B].append(end)

        # Убедимся, что начальная и конечная точки присутствуют в графе
        if start not in graph:
            graph[start] = []
        if end not in graph:
            graph[end] = []

        return graph

    def bfs_path(self,graph, start, end):
        """
        Находит путь от точки start к точке end с использованием поиска в ширину (BFS).

        :param graph: Граф в виде словаря.
        :param start: Начальная точка (x, y, z).
        :param end: Конечная точка (x, y, z).
        :return: Список списков, представляющих путь от start к end.
        """
        if start == end:
            return [start]

        visited = set()
        queue = deque([(start, [start])])

        while queue:
            current, path = queue.popleft()
            if current == end:
                return path
            if current not in visited:
                visited.add(current)
                for neighbor in graph[current]:
                    if neighbor not in visited:
                        queue.append((neighbor, path + [neighbor]))
        
        return []


#test = AutoBendFinder()
#test.get_tube_route()







