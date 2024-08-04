# -*- coding: utf-8 -*-

import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, 
                             QGridLayout, QGroupBox)
from PyQt5.QtGui import QIcon

from resources import *

# Стиль в формате QSS
stylesheet = """
QWidget {
    background-color: #39455e;  /* Фон для всего приложения */
}

QPushButton {
    background-color: #576a91;  /* Фон кнопок */
    color: #d3dced;               /* Текст */
    border: none;               /* Без границ */
    padding: 10px 20px;         /* Отступы внутри кнопок */
    font-size: 16px;            /* Размер шрифта */
    border-radius: 5px;         /* Округлые углы */
    min-width: 100px;           /* Минимальная ширина кнопки */
}

QPushButton:hover {
    background-color: #45a049;  /* Цвет кнопки при наведении */
}

QPushButton:pressed {
    background-color: #388e3c;  /* Цвет кнопки при нажатии */
}

QLabel {
    color: #333;                 /* Цвет текста */
    font-size: 18px;            /* Размер шрифта */
    margin: 10px;               /* Отступы вокруг текста */
}

QFrame {
    border: 1px solid #e0e0e0;  /* Светло-серая граница */
    border-radius: 10px;        /* Округлые углы */
    padding: 10px;              /* Отступы внутри фрейма */
}

QGroupBox {
    border: 1px solid #4CAF50;  /* Зеленая граница для групп */
    border-radius: 10px;        /* Округлые углы */
    padding: 10px;              /* Отступы внутри группы */
    font-size: 14px;            /* Размер шрифта заголовка группы */
    color: #d3dced;             /* Цвет шрифта заголовка группы */
}
"""

class NerpaWindow(QMainWindow):
    def __init__(self) :
        super().__init__()
        self.setWindowTitle('NerpaAI v2.0 QT')
        self.setWindowIcon(QIcon(
            get_resource_path('resources\\pic\\2.ico')
        ))
        #self.setFixedSize(325,575)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        main_layout = QGridLayout(central_widget)

        self.frames = self.initialize_frames()
        for i, frame in enumerate(self.frames.values()):
            main_layout.addWidget(frame)

        self.initialize_buttons()

        self.setStyleSheet(stylesheet)

        self.center_window()

    def initialize_frames(self):
        frames = {}

        frames['assembly'] = QGroupBox('Модуль для работы со сборкой')
        frames['assembly'].setLayout(QVBoxLayout())

        frames['drawing'] = QGroupBox('Модуль для работы с чертежом')
        frames['drawing'].setLayout(QVBoxLayout())

        frames['extra'] = QGroupBox('Дополнительно')
        frames['extra'].setLayout(QVBoxLayout())

        return frames
    
    def initialize_buttons(self):
        button_config = [
            {'text': 'Адаптировать сборку насквозь', 'frame': 'assembly', 
             'command': self.skip, 'enabled': False},
            {'text': 'Адаптировать активную сборку', 'frame': 'assembly', 
             'command': AdaptAssy, 'enabled': True},
            {'text': 'Адаптировать активную деталь', 'frame': 'assembly', 
             'command': AdaltDetail, 'enabled': True},
            {'text': 'Добавить свойства ISO-RGSH', 'frame': 'assembly', 
             'command': self.check_add_prop, 'enabled': True},
            {'text': 'Добавить BOM', 'frame': 'drawing', 
             'command': BOMMaker, 'enabled': True},
            {'text': 'Создать МТО', 'frame': 'extra', 
             'command': MTOMaker, 'enabled': True},
            {'text': 'Установить позиции в активной сборке', 'frame': 'drawing', 
             'command': SetPositions, 'enabled': True},
            {'text': 'Добавить тех. требования', 'frame': 'drawing', 
             'command': TechDemandWindow, 'enabled': True},
            {'text': 'Добавить таблицу гибов', 'frame': 'drawing', 
             'command': BTWindow, 'enabled': True},
            {'text': 'Перевести .CDW чертежи ENG-RUS', 'frame': 'drawing', 
             'command': TranslateCDW, 'enabled': True},
            {'text': 'Редактор словаря', 'frame': 'extra', 
             'command': DictionaryWindow, 'enabled': True},
            {'text': 'Создать PDF', 'frame': 'extra', 
             'command': PDFWindow, 'enabled': True},
        ]

        for config in button_config:
            frame = self.frames[config['frame']]
            button = QPushButton(config['text'])
            button.setEnabled(config['enabled'])
            button.clicked.connect(config['command'])
            frame.layout().addWidget(button)

    def skip(self):
            print("Skip button clicked")

    def check_add_prop(self):
        func = PropertyManager()
        func.check_add_properties()

    def center_window(self):
        # Получаем размеры экрана
        screen_rect = QApplication.desktop().screenGeometry()
        screen_width = screen_rect.width()
        screen_height = screen_rect.height()

        # Получаем размеры окна
        window_width = self.width()
        window_height = self.height()

        # Вычисляем позицию окна для размещения его по центру экрана
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2

        # Устанавливаем позицию окна
        self.move(x, y)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = NerpaWindow()
    window.show()
    sys.exit(app.exec_())

