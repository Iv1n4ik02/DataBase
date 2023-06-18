import sys
from PyQt5.QtWidgets import QApplication, QWidget, QCalendarWidget
from PyQt5.QtGui import QIcon
class CalendarWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle('Календарь')
        self.setGeometry(200, 200, 500, 240)
        self.calendar = QCalendarWidget(self)
        self.calendar.setGeometry(0, 0, 500, 240)
        self.calendar.setGridVisible(True)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = CalendarWindow()
    icon = QIcon('icon.png')
    window.setWindowIcon(icon)
    window.show()
    sys.exit(app.exec_())