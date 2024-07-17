# -*- coding: utf-8 -*-

from PyQt5.QtWidgets import (QApplication, QMainWindow, 
                            QPushButton, QLineEdit, QComboBox, 
                            QWidget, QGridLayout,
                            QVBoxLayout, QGroupBox,
                            QTreeView, QLabel, QMessageBox,
                            QHeaderView)

from PyQt5.QtGui import QStandardItem, QStandardItemModel

import sqlite3
import sys


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('RGSH Dictionary Editor')
        self.setMaximumHeight(400)
        self.setMinimumWidth(400)
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        self.db_mng = DBManager()

        self.init_ui()

    def init_ui(self):
        self.main_layout = QVBoxLayout(self.central_widget)
        self.central_widget.setLayout(self.main_layout)

        self.manage_bar()
        self.find_bar()
        self.treeview_bar()

        self.main_layout.addStretch()

    def manage_bar(self):
        self.manage_box = QGroupBox('Управление')
        self.manage_box.setMaximumHeight(100)
        self.manage_names = (('Добавить', 0, 0),
                              ('Изменить', 0, 1),
                              ('Удалить', 0, 2))
        self.manage_layout = QGridLayout()

        self.manage_buttons = [QPushButton(name[0])
                                for name in self.manage_names]
        for i, button in enumerate(self.manage_buttons):
            self.manage_layout.addWidget(button, self.manage_names[i][1],
                                         self.manage_names[i][2])
            
        self.manage_buttons[0].clicked.connect(self.add_word_button)
        self.manage_buttons[2].clicked.connect(self.delete_row)
            
        self.manage_box.setLayout(self.manage_layout)
        self.central_widget.layout().addWidget(self.manage_box)

    def add_word_button(self):
        self.add_word_window = AddWord()
        self.add_word_window.show()

    def find_bar(self):
        self.find_box = QGroupBox('Найти слово')
        self.find_box.setMaximumHeight(200)
        self.find_layout = QGridLayout()

        self.find_text = QLineEdit(self)
        self.find_layout.addWidget(self.find_text,0,0,1,2)

        self.find_button = QPushButton('Найти')
        self.find_layout.addWidget(self.find_button,1,1)

        self.find_language = QComboBox(self)
        self.find_language.addItems(['Русский', 'Английский'])
        self.find_layout.addWidget(self.find_language,1,0)

        self.find_box.setLayout(self.find_layout)
        self.central_widget.layout().addWidget(self.find_box)

    def treeview_bar(self):
        self.tree_box = QGroupBox('Словарь')
        self.tree_layout = QGridLayout()

        global treeview, model
        treeview = QTreeView()
        model = QStandardItemModel()
        model.setHorizontalHeaderLabels(['Русский', 'Английский'])
        treeview.setIndentation(0)

        dictionary = self.db_mng.get_dictionary()
        for key, value in dictionary.items():
            item = [QStandardItem(key), QStandardItem(value)]
            model.appendRow(item)

        treeview.setModel(model)
        treeview.header().setSectionResizeMode(QHeaderView.Stretch)
        self.tree_layout.addWidget(treeview)
        self.tree_box.setLayout(self.tree_layout)
        self.central_widget.layout().addWidget(self.tree_box)

    def delete_row(self):
        selected_rows = treeview.selectionModel().selectedIndexes()
        if selected_rows:
            index = selected_rows[0]
            item = model.itemFromIndex(index)
            rus_word = item.text()
            self.db_mng.delete_row(rus_word)
            model.removeRow(index.row())
            treeview.update()
            
            
class AddWord(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Добавить пару слов')
        self.setFixedSize(400,110)
        self.db_mng = DBManager()
        self.init_ui()

    def init_ui(self):
        self.wind_layout = QGridLayout()

        self.rus_label = QLabel('Русский',self)
        self.en_label = QLabel('Английский', self)

        self.rus_line = QLineEdit(self)
        self.en_line = QLineEdit(self)

        self.add_button = QPushButton('Добавить')
        self.add_button.clicked.connect(self.add_word)

        self.wind_layout.addWidget(self.rus_label, 0, 0)
        self.wind_layout.addWidget(self.en_label, 0, 1)
        self.wind_layout.addWidget(self.rus_line, 1, 0)
        self.wind_layout.addWidget(self.en_line, 1, 1)
        self.wind_layout.addWidget(self.add_button, 2,0,1,0)

        self.setLayout(self.wind_layout)

    def add_word(self):
        values = (self.rus_line.text(), self.en_line.text())
        for value in values:
            if value == '':
                get_error_msg(
                    'Введите данные в поля для ввода')
                return
        add_word_state = self.db_mng.add_word(values)
        if add_word_state:
            get_info_msg(
                'Пара слов успешно добавлена в базу')
            item = [QStandardItem(values[0]), QStandardItem(values[1])]
            model.appendRow(item)
            treeview.update()
            return
        
        get_error_msg(
            'Одно из слов уже находится в базе.\nПроверьте правильность заполнения')

class DBManager():
    def __init__(self):
        self.db_path = "D:\\GIT base\\NerpaAI\\lib\\DICTIONARY.db"
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()
        self.table_name = 'dictionary'

    def create_table(self):
        self.cursor.execute(f"""
                            CREATE TABLE IF NOT EXISTS
                            {self.table_name}
                            (id INTEGER PRIMARY KEY AUTOINCREMENT,
                            rus_word TEXT,
                            en_word TEXT)
                            """)

    def get_dictionary(self):
        self.cursor.execute(f"""
                            SELECT rus_word, en_word
                            FROM {self.table_name}""")
        dict_list = self.cursor.fetchall()
        keys, values = zip(*dict_list)
        return dict(zip(keys, values))
            
    def get_column_info(self, column_name):
        self.cursor.execute(f"""SELECT {column_name}
                            FROM {self.table_name}""")
        column_info = self.cursor.fetchall()
        info_list = [info[0] for info in column_info]

        return info_list

    def add_word(self, values):
        dictionary = self.get_dictionary()
        for value in values:
            if value in dictionary:
                return False
            if value in dictionary.values():
                return False

        self.cursor.execute(f"""
                            INSERT INTO {self.table_name}
                            (rus_word, en_word)
                            VALUES (?,?)
                               """, (values))
        self.conn.commit()

        return True
    
    def delete_row(self, value):
        delete_query = f"""DELETE from {self.table_name} where rus_word = ?"""
        self.cursor.execute(delete_query, (value, ))
        self.conn.commit()

def get_error_msg(error_text):
    msg_box = QMessageBox()
    msg_box.setIcon(QMessageBox.Critical)
    msg_box.setWindowTitle('Ошибка')
    msg_box.setText(error_text)
    msg_box.exec()

def get_info_msg(info_text):
    msg_box = QMessageBox()
    msg_box.setIcon(QMessageBox.Information)
    msg_box.setWindowTitle('Успех')
    msg_box.setText(info_text)
    msg_box.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    dict_window = MainWindow()
    dict_window.show()
    sys.exit(app.exec_())