from PySide import QtGui
import sys
import requests
import calendar
import time
import datetime
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook
from openpyxl import load_workbook

API_URL = "http://api.openweathermap.org/data/2.5/forecast?q="
API_ID = "&APPID=4c8807c1a33b7cbd1f3f04a2f03a0bf3"

class HttpCodes:
    ok = 200
    bad_request = 400
    unauthorized = 401
    not_found = 404

class HistoricalWeatherProvider:

    def getHistoricalWeatherForCity(self, city_name):
        raise NotImplementedError

    def getAllHistoricalTemperatures(self):
        raise NotImplementedError

class OWMHistoricalWeatherClient(HistoricalWeatherProvider):

    def __init__(self, city_name, end):
        self.city = city_name
        self.end_date = end

    def getHistoricalWeatherForFiveDaysForCity(self):

        data_to_export = {}

        user_end_date = calendar.timegm(time.strptime(self.end_date, '%d.%m.%Y'))

        for every_city in self.city:
            url = API_URL + every_city +"&units=metric" + API_ID
            print(url)
            response = requests.get(url)

            if HttpCodes.ok == response.status_code:
                json_string = response.json()
                list_length = len(json_string['list'])
                weather_list = []
                temp_max_list = []
                temp_min_list = []
                for every_date in range(list_length):

                    if(user_end_date >= json_string['list'][every_date]['dt']):
                        weather_list.append(json_string['list'][every_date]['main']['temp'])
                        temp_max_list.append(json_string['list'][every_date]['main']['temp_max'])
                        temp_min_list.append(json_string['list'][every_date]['main']['temp_min'])
                        end_date = json_string['list'][every_date]['dt_txt']
                        start_date = json_string['list'][0]['dt_txt']
                print("Sucessfully downloaded temperature from service")

            else:
                print(response.status_code)
                return None

            data_to_export.update({every_city: {'Average temperature': sum(weather_list)/list_length,
                                                'Temp_max': max(temp_max_list),
                                                'Temp_min': min(temp_min_list),
                                                'End date': end_date,
                                                'Start date': start_date}})
        print("Data to export: \n", data_to_export)
        return data_to_export

class DocxHistoricalWeatherReporter:

    def __init__(self, report_source):
        self.report_path = report_source

    def generateReportFromWeb(self, weather_data):
        document = Document()

        document.add_heading('Temperature Report', 0)

        if weather_data is not None:
            for every_city in weather_data:
                document.add_heading(every_city, level=1)
                table = document.add_table(rows=2, cols=5)

                temperature_of_city = weather_data[every_city]['Average temperature']
                end_date = weather_data[every_city]['End date']
                start_date = weather_data[every_city]['Start date']
                temp_max = weather_data[every_city]['Temp_max']
                temp_min = weather_data[every_city]['Temp_min']

                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'From'
                hdr_cells[1].text = 'To'
                hdr_cells[2].text = 'Average'
                hdr_cells[3].text = 'Max'
                hdr_cells[4].text = 'Min'

                hdr_data = table.rows[1].cells
                hdr_data[0].text = str(start_date)
                hdr_data[1].text = str(end_date)
                hdr_data[2].text = str(round(temperature_of_city))
                hdr_data[3].text = str(temp_max)
                hdr_data[4].text = str(temp_min)
        else:
            print("Couldn't download temperature from service")
            return None

        document.add_page_break()

        document.save(self.report_path)

        print("Word raport successfully generated from Web Service and Excel")



class Example(QtGui.QWidget):

    def __init__(self):
        super(Example, self).__init__()

        self.initUI()

        self.cityListToExport = []

    def initUI(self):

        self.st = QtGui.QLabel("Miasto:")

        self.textBox = QtGui.QTextEdit(self)
        self.textBox.setMaximumSize(150,30)

        self.buttonAddCity = QtGui.QPushButton('Add City', self)

        self.cityList = QtGui.QListWidget(self)
        self.cityList.readOnly = True

        self.buttonGetFromWeb = QtGui.QPushButton('Get Temperature From Web', self)

        #self.resize(500, 600)
        self.center()
        self.setWindowTitle('Center')

        layout = QtGui.QVBoxLayout()
        layout.addWidget(self.st)
        layout.addWidget(self.textBox)
        layout.addWidget(self.buttonAddCity)
        layout.addWidget(self.cityList)
        layout.addWidget(self.buttonGetFromWeb)

        self.setLayout(layout)

        self.buttonAddCity.clicked.connect(self.btnstateForCityButton)
        self.buttonGetFromWeb.clicked.connect(self.btnstateForWebService)

    def btnstateForCityButton(self):
        if self.buttonAddCity.isChecked():
            print("button pressed")
        else:
            city = self.textBox.toPlainText()
            self.cityList.addItem(city)
            self.cityListToExport.append(city)
            self.textBox.clear()

    def btnstateForWebService(self):
        if self.buttonGetFromWeb.isChecked():
            print("button pressed")
        else:
            weather_from_internet = OWMHistoricalWeatherClient(self.cityListToExport, "23.03.2017")
            temperature_of_cities = weather_from_internet.getHistoricalWeatherForFiveDaysForCity()

            docx = DocxHistoricalWeatherReporter("C:/Users/krzacjan/Desktop/pythonHomeWork/temperature_report_from_web.docx")
            docx.generateReportFromWeb(temperature_of_cities)

    def center(self):
        qr = self.frameGeometry()
        cp = QtGui.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

def main():
    app = QtGui.QApplication(sys.argv)
    ex = Example()
    ex.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()