class PathFinder:
    def __init__(self, curves, start_point, end_point):
        self.curves = curves
        self.start_point = start_point
        self.end_point = end_point
        self.visited = set()
        self.graph = self.build_graph()

    def build_graph(self):
        graph = {}
        for curve in self.curves:
            _, x1, y1, z1, x2, y2, z2 = curve
            point1 = (x1, y1, z1)
            point2 = (x2, y2, z2)
            if point1 not in graph:
                graph[point1] = []
            if point2 not in graph:
                graph[point2] = []
            graph[point1].append(point2)
            graph[point2].append(point1)
        return graph

    def can_find_path(self, current_point, end_point):
        if current_point in self.visited:
            return False
        if current_point == end_point:
            return True
        self.visited.add(current_point)
        for neighbor in self.graph.get(current_point, []):
            if self.can_find_path(neighbor, end_point):
                return True
        return False

    def is_path_possible(self):
        return self.can_find_path(self.start_point, self.end_point)

curves = [
    (True, 0.0, 0.0, 0.0, 0.0, 300.0, 0.0),
    (True, -984.0, 300.0, 0.0, 984.0, 300.0, 0.0),
    (True, 984.0, -500.0, 0.0, 984.0, 300.0, 0.0),
    (True, 984.0, -500.0, 0.0, 1848.0, -500.0, 0.0),
    (True, -984.0, 300.0, 0.0, -984.0, 300.0, 1174.0),
    (True, -984.0, -410.0, 1174.0, -984.0, 300.0, 1174.0),
    (True, -984.0, -410.0, 1174.0, -391.0, -410.0, 1174.0)
]

start_point = (110.3, 300.0, 0.0)
end_point = (1848.0, -500.0, 0.0)

path_finder = PathFinder(curves, start_point, end_point)
print(path_finder.is_path_possible())
