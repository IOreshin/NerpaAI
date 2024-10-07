# -*- coding: utf-8 -*-

import pythoncom
from win32com.client import Dispatch, gencache
from NerpaUtility import KompasAPI


class BendPipeLines(KompasAPI):
    def __init__(self):
        super().__init__()
        self.module5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
        self.app5 = self.module5.KompasObject(
            Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(self.module5.KompasObject.CLSID, pythoncom.IID_IDispatch))

        self.edges_dispatchs = []

        self.get_edges_dispatchs()
        self.get_edges_fillet()

    def get_edges_dispatchs(self):
        Document3D = self.app5.ActiveDocument3D()
        iPart = Document3D.GetPart(self.const.pTop_Part)
        iFeature = iPart.GetFeature()
        iEntityCollection = iFeature.EntityCollection(7)
        for i in range(iEntityCollection.GetCount()):
            iEntity = iEntityCollection.GetByIndex(i)
            iDefinition = iEntity.GetDefinition()
            if iDefinition.IsStraight():
                iEdge = self.app5.TransferInterface(iDefinition, 2, 0)
                self.edges_dispatchs.append(iEdge)

    def get_edges_fillet(self):
        iPart7 = self.get_part_dispatch()
        iAuxGeomContainer = self.module.IAuxiliaryGeomContainer(iPart7)
        iFilletsCurves = iAuxGeomContainer.FilletCurves
        iFilletCurve = iFilletsCurves.Add()
        iFilletCurve.Curve1 = self.edges_dispatchs[0]
        iFilletCurve.Curve2 = self.edges_dispatchs[1]
        iFilletCurve.Radius = 75
        iFilletCurve.TrimCurve1 = False
        iFilletCurve.TrimCurve2 = False
        print(iFilletCurve)
        iFilletCurve.Update()
        print(iFilletCurve.GetCurve1CutPoint())
        print(iFilletCurve.GetCurve2CutPoint())


 


BendPipeLines()