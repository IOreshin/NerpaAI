# -*- coding: utf-8 -*-

from .NerpaUtility import KompasAPI, get_resource_path

from .ConstantsRGSH import SHEET_FORMATS

class MTOMaker(KompasAPI):
    '''
    Класс для создания MTO
    '''
    def __init__(self):
        super().__init__()
        self.name = ' MTO'
        self.style = 2
        self.reports_styles_path = get_resource_path(
            'resources/lib/RGSH_REPORTS.lrt')
        self.get_report()

    def get_report(self):
        '''
        Метод для создания МТО
        '''
        try:
            doc = self.app.ActiveDocument
            if doc.DocumentType != 5:
                self.app.MessageBoxEx('Активный файл не является сборкой',
                                    'Ошибка', 64)
                return
        except Exception:
            self.app.MessageBoxEx('Активный файл не является сборкой',
                                    'Ошибка', 64)
            return
        
        doc_path = doc.PathName

        iPropertyMng = self.module.IPropertyMng(self.app)
        iReport = iPropertyMng.GetReport(doc_path, 0)
        iReport.CurrentStyleIndex = self.style
        iReport.ShowAllObjects = False
        iReport.UseReportFilter = True

        iRepObjFilter = self.module.IReportObjectsFilter(iReport)
        iRepObjFilter.Bodies = True
        iRepObjFilter.Parts = True

        iRepParam = self.module.IReportParam(iReport)
        iRepParam.BuildingType = 0
        iRepParam.UseHyperText = False
        report_filename = str(doc_path[:-4])+str(self.name)+str('.xls')

        iReport.Rebuild()
        iReport.SaveAs(report_filename)

        msg = 'MTO is saved in:'+str(report_filename)
        self.app.MessageBoxEx(msg, 'Успех', 64)



class BOMMaker(KompasAPI):
    '''
    Класс для размещения BOM таблицы на чертеже
    '''
    def __init__(self):
        super().__init__()
        self.place_bom_drw()

    def add_view(self):
        '''
        Метод добавления вида на чертеж
        '''
        doc = self.app.ActiveDocument
        iLayoutSheets = doc.LayoutSheets
        iLayoutSheet = iLayoutSheets.Item(0)
        iSheetFormat = iLayoutSheet.Format
        Format = iSheetFormat.Format
        iKompasDocument2D = self.module.IKompasDocument2D(doc)
        iViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager
        iViews = iViewsAndLayersManager.Views
        self.iView = iViews.Add(1)
        self.iDrawingObject = self.module.IDrawingObject(self.iView)
        self.iView.Current = True
        self.iView.Name = "BOM"

        self.iView.X = SHEET_FORMATS[Format][1]
        self.iView.Y = SHEET_FORMATS[Format][2]
        self.iView.Update()
        self.iDrawingObject.Update()

    def add_title(self):
        '''
        Метод для добавления заголовка таблицы
        '''
        iSymbols2DContainer = self.module.ISymbols2DContainer(self.iView)
        DrawingTables = iSymbols2DContainer.DrawingTables
        self.bom_title_height = 8
        bom_title_width = 225
        bom_title = DrawingTables.Add(1,1,self.bom_title_height, bom_title_width, 0)
        iTable = self.module.ITable(bom_title)
        iTableCell = iTable.Cell(0,0)
        iText = self.module.IText(iTableCell.Text)
        iText.Str = 'BILL OF MATERIALS'
        bom_title.X = -bom_title_width
        bom_title.Y = 0
        bom_title.Update()
        self.iView.Update()
        self.iDrawingObject.Update()

    def get_bom(self, source_path):
        '''
        Метод для размещения BOM на чертеже
        '''
        self.add_view()
        self.add_title()

        symbols_2d_container = self.module.ISymbols2DContainer(self.iView)
        iAssociationTables = symbols_2d_container.AssociationTables
        iAssociationTable = iAssociationTables.Add(source_path, self.const2D.ksRTPropertiesReport)
        iAssociationTable.TablePlaceType = self.const2D.ksTPRightBottom
        iAssociationTable.LayerNumber = 0

        iReport = iAssociationTable.Report
        iReport.CurrentStyleIndex = 4
        iReport.ShowAllObjects = False
        iReport.UseReportFilter = False

        iReportParam = self.module.IReportParam(iReport)
        iReportParam.BuildingType = self.const2D.ksRBChoiceToLevel
        iReportParam.LevelsCount = 1
        iReportParam.UseHyperText = True

        iReportObjectsFilter = self.module.IReportObjectsFilter(iReport)
        iReportObjectsFilter.Bodies = True
        iReportObjectsFilter.Parts = True

        iReportStyle = iReport.ReportStyle(iReport.CurrentStyleIndex)
        RowHeight = iReportStyle.RowHeight
        TitleHeight = iReportStyle.TitleHeight
        iReport.Rebuild()
        iReportTable = self.module.IReportTable(iReport)
        iReportTable.DisplayMode = True
        RowsCount = iReportTable.RowsCount
        TableHeight = RowHeight*RowsCount+TitleHeight

        ColumnsCount = iReportStyle.ColumnsCount
        TableWidth = 0
        for i in range(ColumnsCount):
            Column = iReportStyle.Column(i)
            Width = Column.Width
            TableWidth = TableWidth + Width

        iAssociationTable.X = -(TableWidth)
        iAssociationTable.Y = -(TableHeight)-self.bom_title_height
        iAssociationTable.Rebuild()

    def place_bom_drw(self):
        '''
        Основной метод для проверки коллекции видов в чертеже
        и добавления BOM таблицы
        '''
        doc = self.app.ActiveDocument
        if doc.DocumentType != 1:
            self.app.MessageBoxEx('Активный документ не является чертежом',
                                  'Ошибка', 64)
            return
        
        iKompasDocument2D = self.module.IKompasDocument2D(doc)
        ViewsAndLayersManager = iKompasDocument2D.ViewsAndLayersManager
        iViews = ViewsAndLayersManager.Views

        views_counter = 0
        if iViews.Count > 1:
            while views_counter < iViews.Count:
                iView = iViews.View(views_counter)
                if iView.Name == 'BOM':
                    self.app.MessageBoxEx('В чертеже уже есть вид с BOM',
                                              'Ошибка', 64)
                    return
                views_counter += 1
            
            views_counter = 0
            while views_counter < iViews.Count:
                iView = iViews.View(views_counter)
                if iView.Name not in ['Системный вид', 'BOM']:
                    iAssociationView = self.module.IAssociationView(iView)
                    self.get_bom(iAssociationView.SourceFileName)
                    self.app.MessageBoxEx('BOM таблица добавлена в чертеж',
                                          'Успех', 64)
                    return
                views_counter += 1
                
        else:
            self.app.MessageBoxEx('Добавьте ассоциативный вид в чертеж',
                                  'Ошибка', 64)

