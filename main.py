
import requests
import calendar
import time
from docx import Document
from docx.shared import Inches
from openpyxl import Workbook
from openpyxl import load_workbook

API_URL = "http://api.openweathermap.org/data/2.5/weather/?q="    #lub weather zamiast forecast
API_ID = "?id=524901&APPID=4c8807c1a33b7cbd1f3f04a2f03a0bf3"

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

    def getHistoricalWeatherForCity(self):

        for every_city in self.city:
            url = API_URL + every_city  + API_ID

            #url = API_URL + every_city + API_ID + "&type=hour&start=" + str(start_time_hour) + "&end=" + str(
            #   end_time_hour)
            response = requests.get(url)

            if HttpCodes.ok == response.status_code:
                json_string = response.json()
                self.temperature_of_city.update({every_city: {'Temperature': json_string['main']["temp"]}})

            else:
                print(response.status_code)
                return None

        return self.temperature_of_city

class XlsxHistoricalWeatherReader(HistoricalWeatherProvider):

    def __init__(self, city_name, start, end):
        self.city = city_name
        self.temperature_of_city = {}
        self.start_date = start
        self.end_date = end

    def getHistoricalWeatherForCity(self):
        wb = load_workbook('temperature.xlsx')
        ws = wb.active
        print(wb.get_sheet_names())
        for row in ws.iter_rows(min_row=1, max_col=4, max_row=4):
            for cell in row:
                tuple(ws.rows)
                self.temperature_of_city = cell.value




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

                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'From'
                hdr_cells[1].text = 'To'
                hdr_cells[2].text = 'Average'
                hdr_cells[3].text = 'Max'
                hdr_cells[4].text = 'Min'

                hdr_data = table.rows[1].cells
                hdr_data[0].text = "od"
                hdr_data[1].text = "do"
                hdr_data[2].text = str(temperature_of_city)
                hdr_data[3].text = "Max"
                hdr_data[4].text = "Min"
        else:
            print("Couldn't download temperature from service")
            return None

        document.add_page_break()

        document.save(self.report_path)

        print("Word raport successfully generated from Web")

    def generateReportFromExcel(self, weather_data):
            document = Document()

            document.add_heading('Temperature Report', 0)

            if weather_data is not None:
                for every_city in weather_data:
                    document.add_heading(every_city, level=1)
                    table = document.add_table(rows=2, cols=5)

                    temperature_of_city = weather_data[every_city]['Temperature']

                    hdr_cells = table.rows[0].cells
                    hdr_cells[0].text = 'From'
                    hdr_cells[1].text = 'To'
                    hdr_cells[2].text = 'Average'
                    hdr_cells[3].text = 'Max'
                    hdr_cells[4].text = 'Min'

                    hdr_data = table.rows[1].cells
                    hdr_data[0].text = "od"
                    hdr_data[1].text = "do"
                    hdr_data[2].text = str(temperature_of_city)
                    hdr_data[3].text = "Max"
                    hdr_data[4].text = "Min"
            else:
                print("Couldn't parse data from Excel")
                return None

            document.add_page_break()

            document.save(self.report_path)

            print("Word raport successfully generated from Excel")


def main():
    weather_from_internet = OWMHistoricalWeatherClient({"Wroclaw", "Berlin"}, "01.12.2016", "20.12.2016")
    temperature_of_cities = weather_from_internet.getHistoricalWeatherForCity()
    print(temperature_of_cities)

    docx = DocxHistoricalWeatherReporter("C:/Users/krzacjan/Desktop/pythonHomeWork/temperature_report_from_web.docx")
    docx.generateReportFromWeb(temperature_of_cities)

    #weather_from_excel =  XlsxHistoricalWeatherReader({"Wroclaw", "Berlin"}, "01.12.2016", "20.12.2016")
    #temperature_of_cities_from_excel = weather_from_excel.getHistoricalWeatherForCity()

    #docx = DocxHistoricalWeatherReporter("C:/Users/krzacjan/Desktop/pythonHomeWork/temperature_report_from_excel.docx")
    #docx.generateReportFromExcel(temperature_of_cities_from_excel)

if __name__ == "__main__":
    main()



