import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QPushButton
from PyQt5.QtCore import Qt

# Стиль в формате QSS
stylesheet = """
QWidget {
    background-color: #f5f5f5;  /* Светлый фон для всего приложения */
}

QPushButton {
    background-color: #4CAF50;  /* Зеленый фон кнопок */
    color: white;               /* Белый текст */
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
    font-size: 16px;            /* Размер шрифта заголовка группы */
    font-weight: bold;          /* Полужирный текст заголовка */
}
"""

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Стильное Приложение")
        self.setFixedSize(600, 400)
        
        # Основной виджет и макет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Заголовок
        title = QLabel("Добро пожаловать в стильное приложение!")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)

        # Кнопки
        button1 = QPushButton("Кнопка 1")
        button2 = QPushButton("Кнопка 2")
        button3 = QPushButton("Кнопка 3")

        layout.addWidget(button1)
        layout.addWidget(button2)
        layout.addWidget(button3)

        # Применение стиля
        self.setStyleSheet(stylesheet)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
