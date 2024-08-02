# -*- coding: utf-8 -*-

from .NerpaUtility import KompasAPI

class AutoTubeCdw(KompasAPI):
    def __init__(self):
        super().__init__()
        self.iDocuments = self.app.Documents
        #self.get_doc_3d()
        self.create_cdw()


    def get_doc_3d(self):
        self.assy_doc = self.app.ActiveDocument
        self.doc3D = self.module.IKompasDocument3D(self.assy_doc)
        self.doc3D_path = self.assy_doc.PathName
        self.doc3D_object = self.module.IModelObject(self.doc3D)
        bodies = self.module.IFeature7(self.doc3D.TopPart).ResultBodies
        print(bodies)
        

    def _get_3d(self):
        self.doc = self.iDocuments.Item(1)
        self.doc3D = self.module.IKompasDocument3D(self.doc)
       

    def create_cdw(self):
        #self.cdw_doc = self.iDocuments.Add(1, True)
        #self.cdw_doc.Active = True
        self.cdw_doc = self.app.ActiveDocument

        self.doc2D = self.module.IKompasDocument2D(self.cdw_doc)
        self.doc2D1 = self.module.IKompasDocument2D1(self.doc2D)
        
        views_mng = self.doc2D.ViewsAndLayersManager
        views = views_mng.Views
        #view = views_mng.Add(2)
        view = views.ViewByNumber(1)
        
        
        view.X = 0
        view.Y = 0

        asso_view = self.module.IAssociationView(view)
        #asso_view.SourceFileName = self.doc3D_path
        #asso_view.ProjectionName = "#Спереди"
        self._get_3d()

        asso_objects = asso_view.GetProjectionObjects(self.doc3D)
        print(asso_objects[0])
        asso_view.SetProjectionObjects(asso_objects)

        view.Update()
        self.doc2D1.RebuildDocument()


AutoTubeCdw()