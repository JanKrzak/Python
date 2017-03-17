
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

    def __init__(self, city_name, start, end):
        self.city = city_name
        self.temperature_of_city = {}
        self.start_date = start
        self.end_date = end

    def getHistoricalWeatherForFiveDaysForCity(self):

        for every_city in self.city:
            url = API_URL + every_city +"&units=metric" + API_ID
            print(url)
            response = requests.get(url)

            if HttpCodes.ok == response.status_code:
                json_string = response.json()
                self.temperature_of_city.update({every_city: {'Temperature': {json_string['list'][0]['main']['temp']}, 'Date': {json_string['list'][0]['dt']},
                                                              'temp_max': {json_string['list'][0]['main']['temp_max']}, "temp_min": {json_string['list'][0]['main']['temp_min']}}})
                print("Sucessfully downloaded temperature from service")
            else:
                print(response.status_code)
                return None

        print('Json from web: \n',self.temperature_of_city)
        return self.temperature_of_city

class XlsxHistoricalWeatherReader(HistoricalWeatherProvider):

    def __init__(self, city_name, start, end):
        self.city = city_name
        self.temperature_of_city = []
        self.city_list = []
        self.date_list = []
        self.temperature = []

    def getHistoricalWeatherForCity(self):
        wb = load_workbook('temperature.xlsx')
        ws = wb.active

        for col in ws.iter_cols(min_row=1, max_col=10, max_row=10):
            for cell in col:
                if (cell.value is not None):
                    self.temperature_of_city.append(cell.value)
                else:
                     break

        for row in ws.iter_rows(min_row=2, max_col=1, max_row=10):
            for cell in row:
                if (cell.value is not None):
                    self.date_list.append(cell.value)
                else:
                    break

        for row in ws.iter_rows(max_row=1, min_col=2, max_col=10):
            for cell in row:
                if (cell.value is not None):
                    self.city_list.append(cell.value)
                else:
                    break

        print('List from excel: \n', self.temperature_of_city)
        return self.temperature_of_city


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

                temperature_of_city = weather_data[every_city]['Temperature']
                date = weather_data[every_city]['Date']
                temp_max = weather_data[every_city]['temp_max']
                temp_min = weather_data[every_city]['temp_min']

                for time in date:
                    date_to_int = time
                    break
                date_of_weather = datetime.datetime.fromtimestamp(date_to_int)

                for temperature in temperature_of_city:
                    temperature_to_string = temperature
                    break

                for temp in temp_max:
                    temperature_max_to_string = temp
                    break

                for temp in temp_min:
                    temperature_min_to_string = temp
                    break


                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'From'
                hdr_cells[1].text = 'To'
                hdr_cells[2].text = 'Average'
                hdr_cells[3].text = 'Max'
                hdr_cells[4].text = 'Min'

                hdr_data = table.rows[1].cells
                hdr_data[0].text = str(date_of_weather)
                hdr_data[1].text = "do"
                hdr_data[2].text = str(temperature_to_string)
                hdr_data[3].text = str(temperature_max_to_string)
                hdr_data[4].text = str(temperature_min_to_string)
        else:
            print("Couldn't download temperature from service")
            return None

        document.add_page_break()

        document.save(self.report_path)

        print("Word raport successfully generated from Web Service")

    def generateReportFromExcel(self, city_list, date_list):
            document = Document()

            document.add_heading('Temperature Report', 0)

            if city_list is not None:
                for every_city in city_list:
                    document.add_heading(every_city, level=1)
                    table = document.add_table(rows=2, cols=5)

                    hdr_cells = table.rows[0].cells
                    hdr_cells[0].text = 'From'
                    hdr_cells[1].text = 'To'
                    hdr_cells[2].text = 'Average'
                    hdr_cells[3].text = 'Max'
                    hdr_cells[4].text = 'Min'

                    hdr_data = table.rows[1].cells
                    hdr_data[0].text = date_list[0]
                    hdr_data[1].text = date_list[-1]
                    #hdr_data[2].text = str(temperature_of_city)
                    #hdr_data[3].text = "Max"
                    #hdr_data[4].text = "Min"
            else:
                print("Couldn't parse data from Excel")
                return None

            document.add_page_break()

            document.save(self.report_path)

            print("Word raport successfully generated from Excel")


def main():

    city_list = []

    weather_from_internet = OWMHistoricalWeatherClient({"Wroclaw", "Berlin"}, "01.12.2016", "20.12.2016")
    temperature_of_cities = weather_from_internet.getHistoricalWeatherForFiveDaysForCity()
    print(temperature_of_cities)

    docx = DocxHistoricalWeatherReporter("C:/Users/krzacjan/Desktop/pythonHomeWork/temperature_report_from_web.docx")
    docx.generateReportFromWeb(temperature_of_cities)

    weather_from_excel =  XlsxHistoricalWeatherReader({"Wroclaw", "Berlin"}, "01.12.2016", "20.12.2016")
    temperature_of_cities_from_excel = weather_from_excel.getHistoricalWeatherForCity()

    docx = DocxHistoricalWeatherReporter("C:/Users/krzacjan/Desktop/pythonHomeWork/temperature_report_from_excel.docx")
    docx.generateReportFromExcel(weather_from_excel.city_list, weather_from_excel.date_list)

if __name__ == "__main__":
    main()



