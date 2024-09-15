# -*- coding: utf-8 -*-
from NerpaUtility import KompasAPI

class ChangeStampName(KompasAPI):
    def __init__(self):
        super().__init__()
        self.iDocuments = self.app.Documents
        self.change_stamp_name()


    def change_stamp_name(self):
        for i in range(self.iDocuments.Count):
            doc = self.iDocuments.Item(i)
            if doc.DocumentType == 1:
                iLayoutSheets = doc.LayoutSheets
                iLayoutSheet = iLayoutSheets.Item(0)
                iStamp = iLayoutSheet.Stamp
                text1 = iStamp.Text(250)
                text1.Str = 'ANNA P.'

                text2 = iStamp.Text(260)
                text2.Str = 'ALEX L.'

                text2 = iStamp.Text(270)
                text2.Str = 'ANTON S.'

                iStamp.Update()


ChangeStampName()
