# -*- coding: utf-8 -*-
from .NerpaUtility import KompasAPI, read_json
from .WindowModule import Window
import tkinter as tk
from tkinter import ttk
import datetime
import getpass

class StampWindow(Window):
    def __init__(self):
        super().__init__()
        self.user = self.get_user_name()

        self.init_ui()

    def get_user_name(self):
        username = getpass.getuser()

        users = (('oreshin_ie', 'IVAN O.'),
                ('bezginov_vo', 'VLADIMIR B.'),
                ('meshkov_ak', 'ALEX M.'),
                ('grigoreva_oy', 'OLGA G.'))

        for user in users:
            if username in user[0]:
                return user[1]
        return ''

    def init_ui(self):
        self.stamp_root = tk.Tk()
        self.stamp_root.title(self.window_name+': Заполнить штамп')
        self.stamp_root.iconbitmap(self.pic_path)
        self.stamp_root.resizable(False, False)
        self.stamp_root.attributes("-topmost", True)

        self.main_frame = tk.Frame(self.stamp_root)

        self.stamp_frame = ttk.LabelFrame(self.stamp_root,
                                          borderwidth=5,
                                          relief='solid',
                                          text = 'Штамп')
        self.stamp_frame.grid(row = 0, column = 0,
                               pady = 5, padx = 5,
                               sticky = 'nsew')

        entry_names = (('REV', 0, 0),
                        ('DATE', 0, 1),
                        ('DESCRIPTION', 0, 2),
                        ('DESIGNED', 0, 3),
                        ('CHECKED 1', 0, 4),
                        ('CHECKED 2', 2, 0),
                        ('APPROVED', 2, 1),
                        ('MATERIAL', 2, 2),
                        ('REV', 2, 3),
                        ('NOC', 2, 4))

        entry_params = (('REV', 1, 0, '00', 5),
                            ('DATE', 1, 1, datetime.date.today().strftime("%d.%m.%Y"), 10),
                            ('DESCRIPTION', 1, 2, 'ISSUED FOR REVIEW', 20),
                            ('DESIGNED', 1, 3, self.user, 10),
                            ('CHECKED 1', 1, 4, 'ANNA P.', 10),
                            ('CHECKED 2', 3, 0, 'ALEX L.', 10),
                            ('APPROVED', 3, 1, 'ANTON S.', 10),
                            ('MATERIAL', 3, 2, 'N/A', 10),
                            ('REV', 3, 3, '00', 5),
                            ('NOC', 3, 4, '0', 5))

        self.entries = [ttk.Entry(self.stamp_frame, width=entry_param[4])
                   for entry_param in entry_params]
        entry_labels = [tk.Label(self.stamp_frame, text=entry_name[0])
                       for entry_name in entry_names]

        for i, entry_label in enumerate(entry_labels):
            entry_label.grid(row = entry_names[i][1],
                            column = entry_names[i][2],
                            pady = 2, padx = 2,
                            sticky = 'nsew')

        for i, entry in enumerate(self.entries):
            entry.grid(row = entry_params[i][1],
                        column = entry_params[i][2],
                        pady = 5, padx = 5,
                        sticky = 'nsew')
            entry.insert(0, entry_params[i][3])

        self.manage_frame = ttk.LabelFrame(self.stamp_root,
                                          borderwidth=5,
                                          relief='solid',
                                          text = 'Изменить')
        self.manage_frame.grid(row = 1, column = 0,
                               pady = 5, padx = 5,
                               sticky = 'nsew')

        self.create_button(ttk, self.manage_frame,
                    'В активном документе', self.change_active_doc,
                    22, 'normal', 0, 0)
        self.create_button(ttk, self.manage_frame,
                    'Во всех документах', self.change_all_docs,
                    22, 'normal', 0, 1)
        self.create_button(ttk, self.manage_frame,
                    'Справка', StampHelpWindow,
                    22, 'normal', 0, 2)

        self.stamp_root.update_idletasks()
        w, h = self.get_center_window(self.stamp_root)
        self.stamp_root.geometry('+{}+{}'.format(w, h))

        self.stamp_root.mainloop()

    def get_entries_values(self):
        entry_input = []
        for entry in self.entries:
            text = entry.get()
            entry_input.append(text)

        return entry_input

    def change_active_doc(self):
        entry_input = self.get_entries_values()
        ChangeStampName(entry_input).change_active_doc_stamp()

    def change_all_docs(self):
        entry_input = self.get_entries_values()
        ChangeStampName(entry_input).change_all_docs_stamp()

class ChangeStampName(KompasAPI):
    def __init__(self, entry_input):
        super().__init__()
        self.iDocuments = self.app.Documents
        self.entry_input = entry_input
        self.stamp_IDs = (210, #REV
                          220, #DATE
                          230, #DESCRIPTION
                          240, #DESIGNED
                          250, #CHECKED 1
                          260, #CHECKED 2
                          270, #APPROVED
                          3, #MATERIAL
                          201, #REV MAIN
                          200, #NOC
                          )

    def change_stamp(self, doc_dispatch):
        iLayoutSheets = doc_dispatch.LayoutSheets
        iLayoutSheet = iLayoutSheets.Item(0)
        iStamp = iLayoutSheet.Stamp
        try:
            for i, ID in enumerate(self.stamp_IDs):
                iText = iStamp.Text(ID)
                if self.entry_input[i] not in [None, '', ' ']:
                    iText.Str = self.entry_input[i]
                    iStamp.Update()
            self.app.MessageBoxEx('Штамп успешно изменен в документе {}'.format(doc_dispatch.Name),
                                  '', 64)
        except Exception as e:
            self.app.MessageBoxEx('Произошла ошибка при изменении штампа:{}'.format(e),
                                  'Ошибка', 64)

    def change_all_docs_stamp(self):
        for i in range(self.iDocuments.Count):
            doc = self.iDocuments.Item(i)
            if doc and doc.DocumentType == 1:
                self.change_stamp(doc)

    def change_active_doc_stamp(self):
        doc = self.app.ActiveDocument
        if doc and doc.DocumentType == 1:
            self.change_stamp(doc)

class StampHelpWindow(Window):
    def __init__(self):
        super().__init__()
        self.get_help_window()

    def get_help_window(self):
        help_root = tk.Tk()
        help_root.title(self.window_name+' : Справка')
        help_root.iconbitmap(self.pic_path)
        help_root.resizable(False, False)
        help_root.attributes("-topmost", True)

        help_frame = ttk.LabelFrame(help_root, borderwidth = 5, relief = 'solid', text = 'Справка')
        help_frame.grid(row = 0, column = 0, pady = 5, padx = 5, sticky = 'nsew')

        HelpPages = read_json('/lib/HELPPAGES.json' )
        HelpText = ''
        for item, text in HelpPages.items():
            if item == 'CHANGE STAMP HELP':
                for stroka in text:
                    HelpText = HelpText+stroka+'\n'

        help_label = tk.Label(help_frame, text = HelpText, state=["normal"])
        help_label.grid(row = 0, column = 0, sticky = 'e', pady = 2, padx = 5)

        help_root.update_idletasks()
        w,h = self.get_center_window(help_root)
        help_root.geometry('+{}+{}'.format(w,h))
        help_root.mainloop()