import sys
import sqlite3
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout
from PyQt5.QtGui import QIcon

class Window_R(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('График')
        self.setGeometry(400, 400, 500, 500)
        plt.title('Прибыль')
        plt.xlabel('Время')
        plt.ylabel('Деньги')
        # создаем виджет для графика
        graph_widget = QWidget(self)
        graph_widget.setLayout(QVBoxLayout(graph_widget))
        graph_widget.layout().setContentsMargins(0, 0, 0, 0)
        self.setCentralWidget(graph_widget)
        # подключаемся к базе данных SQL
        conn = sqlite3.connect('DataBase.db')
        cur = conn.cursor()

        # выполняем запрос к базе данных для получения данных для графика
        cur.execute("SELECT O_DataEnd, O_Prise FROM Orde WHERE O_state == 1;")
        data = cur.fetchall()

        # закрываем соединение с базой данных
        conn.close()

        # создаем списки для дат и значений
        dates = [row[0] for row in data]
        values = [row[1] for row in data]

        # создаем график
        plt.plot(dates, values, label ="Mark", marker ="o")
        plt.xticks(rotation=45, ha='right')

        # добавляем график на виджет
        graph_widget.layout().addWidget(plt.gcf().canvas)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Window_R()
    icon = QIcon('icon.png')
    window.setWindowIcon(icon)
    window.show()
    sys.exit(app.exec_())
