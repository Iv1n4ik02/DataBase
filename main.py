from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QMessageBox, QLabel, QVBoxLayout, QProgressBar, QWidget
from C import Window_C
from O import Window_O
from Z import Window_Z
from CT1 import Window_CT
from W import Window_W1
from S import Window_S
from P import Window_P
from Cal import CalendarWindow
from R import Window_R
from Weather import WeatherApp
from word import word
from PyQt5.QtGui import QIcon

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow,self).__init__(parent)
        self.second_window1 = None
        self.second_window2 = None  
        self.second_window3 = None  
        self.second_window4 = None  
        self.second_window5 = None  
        self.second_window6 = None  
        self.second_window7 = None  
        self.second_window8 = None    
        self.setup_ui()

    def setup_ui(self):
        self.setWindowTitle('Менеджер')
        self.setGeometry(400, 300, 340, 400)

        darkModeBtn = QtWidgets.QPushButton('Тема', self)
        darkModeBtn.clicked.connect(self.toggleDarkMode)
        darkModeBtn.setStyleSheet("font-size: 20px")
        darkModeBtn.resize(darkModeBtn.sizeHint())
        darkModeBtn.move(15,360)
        Btn = QtWidgets.QPushButton('Погода', self)
        Btn.clicked.connect(self.show_second_window10)
        Btn.setStyleSheet("font-size: 20px")
        Btn.resize(Btn.sizeHint())
        Btn.setFixedWidth(120)
        Btn.move(115,360)
        self.label1 = QtWidgets.QLabel(self)
        self.label1.move(20,180)
        self.label2 = QtWidgets.QLabel(self)
        self.label2.move(9900,9900)
        self.label2.setText("тема1")
        self.label0 = QtWidgets.QLabel(self)
        self.label0.move(20,150)
        self.label3 = QtWidgets.QLabel(self)
        self.label3.move(245,210)

        # Кнопка для открытия второго окна
        btn1 = QtWidgets.QPushButton('Выход', self)
        btn1.setStyleSheet("font-size: 20px")
        btn1.resize(btn1.sizeHint())
        btn1.move(240, 360)
        btn1.clicked.connect(self.close) 
        btn1 = QtWidgets.QPushButton('Клиенты', self)
        btn1.setStyleSheet("font-size: 20px")
        btn1.resize(btn1.sizeHint())
        btn1.move(15, 5)
        btn1.clicked.connect(self.show_second_window1)
        btn2 = QtWidgets.QPushButton('Сотрудники', self)
        btn2.setStyleSheet("font-size: 20px")
        btn2.resize(btn2.sizeHint())
        btn2.move(215, 5)
        btn2.clicked.connect(self.show_second_window5)
        btn3 = QtWidgets.QPushButton('Техника', self)
        btn3.setStyleSheet("font-size: 20px")
        btn3.resize(btn3.sizeHint())
        btn3.move(115, 5)
        btn3.clicked.connect(self.show_second_window4)
        btn4 = QtWidgets.QPushButton('Услуги', self)
        btn4.setStyleSheet("font-size: 20px")
        btn4.resize(btn4.sizeHint())
        btn4.move(15, 55)
        btn4.clicked.connect(self.show_second_window6)
        btn5 = QtWidgets.QPushButton('Склад', self)
        btn5.setStyleSheet("font-size: 20px")
        btn5.resize(btn5.sizeHint())
        btn5.move(115, 55)
        btn5.clicked.connect(self.show_second_window7)
        btn7 = QtWidgets.QPushButton('Заказы', self)
        btn7.setStyleSheet("font-size: 20px")
        btn7.resize(btn7.sizeHint())
        btn7.setFixedWidth(115)
        btn7.move(215, 55)
        btn7.clicked.connect(self.show_second_window2)
        btn6 = QtWidgets.QPushButton('Чеки', self)
        btn6.setStyleSheet("font-size: 20px")
        btn6.resize(btn6.sizeHint())
        btn6.move(15, 105)
        btn6.clicked.connect(self.show_second_window3)
        btn9 = QtWidgets.QPushButton('Календарь', self)
        btn9.setStyleSheet("font-size: 20px")
        btn9.setFixedWidth(115)
        btn9.resize(btn9.sizeHint())
        btn9.move(215, 105)
        btn9.clicked.connect(self.show_second_window8)
        btn8 = QtWidgets.QPushButton('Отчёт', self)
        btn8.setStyleSheet("font-size: 20px")
        btn8.resize(btn8.sizeHint())
        btn8.move(115, 105)
        btn8.clicked.connect(self.show_second_window9)
        btn99 = QtWidgets.QPushButton('Отчёты', self)
        btn99.setStyleSheet("font-size: 20px")
        btn99.resize(btn99.sizeHint())
        btn99.move(10, 145)
        btn99.setFixedWidth(320)
        btn99.clicked.connect(word)
   
    def show_second_window1(self):
        self.second_window1 = Window_C(self)
        self.second_window1.show()
    def show_second_window2(self):
        self.second_window2 = Window_O(self)
        self.second_window2.show()
    def show_second_window3(self):
        self.second_window3 = Window_Z(self)
        self.second_window3.show()
    def show_second_window4(self):
        self.second_window4 = Window_CT(self)
        self.second_window4.show()
    def show_second_window5(self):
        self.second_window5 = Window_W1(self)
        self.second_window5.show()
    def show_second_window6(self):
        self.second_window6 = Window_S(self)
        self.second_window6.show()
    def show_second_window7(self):
        self.second_window7 = Window_P(self)
        self.second_window7.show()
    def show_second_window8(self):
        self.second_window8 = CalendarWindow()
        self.second_window8.show()
    def show_second_window9(self):
        self.second_window9 = Window_R()
        self.second_window9.show()
    def show_second_window10(self):
        self.second_window10 = WeatherApp()
        self.second_window10.show()

    def toggleDarkMode(self):
        i = self.label2.text()
        if i == "тема1":
            self.label2.setText("тема2")
            self.setStyleSheet('background-color: #222; color: #fff;')
        else:
            self.label2.setText("тема1") 
            self.setStyleSheet('background-color: #fff; color: #000;')
    def keyPressEvent(self, event):
        if event.key() == Qt.Key_F1:
            QMessageBox.information(self, "Разработчик", "Иванов Иван 19 пкс-2")
if __name__ == '__main__':
    app = QtWidgets.QApplication([])
    window = MainWindow()
    icon = QIcon('icon.png')
    window.setWindowIcon(icon)
    window.show()
    app.exec_()
