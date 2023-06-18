import requests
import json
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QComboBox, QVBoxLayout, QHBoxLayout, QLineEdit
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon

class WeatherApp(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.city_label = QLabel("Город:", self)
        self.city_combo = QComboBox(self)
        self.city_combo.addItems(["Kazanskoye","Zykovo", "Samara", "Moscow", "Saint Petersburg","Kursk"])
        self.result_label = QLabel("Погода:", self)
        self.result_edit = QLineEdit(self)
        self.result_edit.setReadOnly(True)

        # создаем вертикальный макет для виджета
        self.layout = QVBoxLayout(self)
        
        # создаем горизонтальный макет для QLabel и QComboBox
        self.h_layout = QHBoxLayout()
        self.h_layout.addWidget(self.city_label)
        self.h_layout.addWidget(self.city_combo)
        
        # добавляем горизонтальный макет в вертикальный макет
        self.layout.addLayout(self.h_layout)

        # добавляем QLabel и QLineEdit в вертикальный макет
        self.layout.addWidget(self.result_label)
        self.layout.addWidget(self.result_edit)
        
        # устанавливаем вертикальный макет в качестве основного для виджета
        self.setLayout(self.layout)

        # добавляем обработчик изменения выбранного элемента в выпадающем списке
        self.city_combo.currentIndexChanged.connect(self.get_weather)

        # получаем данные о погоде при запуске приложения для первого города в списке
        self.get_weather()

    def get_weather(self):
        city = self.city_combo.currentText()
        url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&units=metric&appid=eacdcdd8d1ec0aa9b2ea6ff721b72279"

        response = requests.get(url)

        if response.status_code == 200:
            weather_data = json.loads(response.text)

            temperature = weather_data["main"]["temp"]
            feels_like = weather_data["main"]["feels_like"]
            weather = weather_data["weather"][0]["description"]

            weather_string = f"{weather}. {temperature} °C."

            self.result_edit.setText(weather_string)
        else:
            self.result_edit.setText("Error while getting weather data")
            
if __name__ == '__main__':
    app = QApplication([])
    weather_app = WeatherApp()
    icon = QIcon('icon.png')
    weather_app.setWindowIcon(icon)
    weather_app.show()
    app.exec_() 