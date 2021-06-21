from bs4 import BeautifulSoup
import requests, json
import openpyxl
from types import SimpleNamespace
import ssl
import certifi
import geopy.geocoders
ctx = ssl.create_default_context(cafile=certifi.where())
geopy.geocoders.options.default_ssl_context = ctx
from geopy.geocoders import Nominatim
geolocator = Nominatim(user_agent="MyGeoCoder")
import string, json
import xlsxwriter
from datetime import datetime

now = datetime.now()


class OBJ:
    Title = ""
    Address = ""
    FullAddress = ""
    Suburb = ""
    State = ""
    City = ""
    Country = ""
    Postcode = ""
    Latitude = ""
    Longitude = ""

class City:
    Id = 0
    Name = ''

listOBJ = []
listError = []

listCities = []

# for x in range(3, 10):
#     url = "http://www.matahari.co.id/store-locator-ajax"
#     print(url)
#     payload = {
#         '_token': '9KzVu795HHsAiVTF00lgNYMEnqJnxv3OhwEUmRWP',
#         'type': 'province',
#         'id': str(x)
#     }
#
headers = {'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
               'Cookie': '_ga=GA1.3.891129021.1623342189; XSRF-TOKEN=eyJpdiI6ImhONWdkcHVEZVdMQVNtSzlURG5zdWc9PSIsInZhbHVlIjoiRDlMNzhZXC9CTEI1bjBYdm9FODk3NFpYSml6R2NtSU5RdUVxS2kzVUNjQkJqRUJHdENUbUJHTDhcL05qaTlBUzBBIiwibWFjIjoiMDkyZTA0Zjk5ZjdiNGRkZWUyZTZmZDQwMjhjNWJkNDczODhhMTAxMmQzZmQ1Zjc5YWUxYmIzYzFkZTdjYWVkZSJ9; laravel_session=eyJpdiI6Ik9ad0Y5Rm9jMEJmUHdDSmVUMHM3eUE9PSIsInZhbHVlIjoiMndieURibUp4aUZGYys0MnBiQWVtNU9UYUtZV3RXNHFza1hzN2FJMW1FN1gwZjEwYnFFSHUzaGdtTTJiY09ZMyIsIm1hYyI6IjM1ZDkxZmE2ZjkyNTFlNmRmNmU4YzkzY2UwZTM5ZGZhZWQ2YmNlMmMwMTg3M2RlNDc5NTI0NGFlY2IyNDZjNzYifQ==; _gid=GA1.3.1697937999.1623431514; _gat_gtag_UA_121306489_1=1',
               'Host': 'www.matahari.co.id',
               'Origin': 'http: // www.matahari.co.id',
               'Referer': 'http: // www.matahari.co.id / en / store - locator',
               'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36',
               'X-Requested-With': 'XMLHttpRequest'}
#
#     res = requests.request("POST", url, headers=headers, data=payload)
#
#     if res.status_code == 200:
#         output = res.json()
#
#         soup = BeautifulSoup(output['citiesView'], "html.parser")
#
#         output = soup.find('select', {"name": "city"}).find_all('option')
#
#         for option in output:
#             listCities.append({"Id": x, "Name": option['value']})





listCities = [{"Id": 3, "Name": "Bandung"}, {"Id": 3, "Name": "Bangkalan"}, {"Id": 3, "Name": "Banyumas"}, {"Id": 3, "Name": "Batu"}, {"Id": 3, "Name": "Cilegon"}, {"Id": 3, "Name": "Cirebon"}, {"Id": 3, "Name": "Gresik"}, {"Id": 3, "Name": "Jember"}, {"Id": 3, "Name": "Karawang"}, {"Id": 3, "Name": "Kediri"}, {"Id": 3, "Name": "Klaten"}, {"Id": 3, "Name": "Madiun"}, {"Id": 3, "Name": "Magelang"}, {"Id": 3, "Name": "Malang"}, {"Id": 3, "Name": "Mojokerto"}, {"Id": 3, "Name": "Pekalongan"}, {"Id": 3, "Name": "Semarang"}, {"Id": 3, "Name": "Serang"}, {"Id": 3, "Name": "Sidoarjo"}, {"Id": 3, "Name": "Sleman"}, {"Id": 3, "Name": "Sleman Yogyakarta"}, {"Id": 3, "Name": "Sukabumi"}, {"Id": 3, "Name": "Sukoharjo Solo"}, {"Id": 3, "Name": "Surabaya"}, {"Id": 3, "Name": "Surakarta"}, {"Id": 3, "Name": "Tasikmalaya"}, {"Id": 3, "Name": "Tegal"}, {"Id": 3, "Name": "Yogyakarta"}, {"Id": 4, "Name": "Jabodetabek"}, {"Id": 4, "Name": "Tangerang"}, {"Id": 5, "Name": "Balikpapan"}, {"Id": 5, "Name": "Banjar Baru"}, {"Id": 5, "Name": "Banjarmasin"}, {"Id": 5, "Name": "Ketapang"}, {"Id": 5, "Name": "Kotawaringin Timur"}, {"Id": 5, "Name": "Palangkaraya"}, {"Id": 5, "Name": "Pontianak"}, {"Id": 5, "Name": "Samarinda"}, {"Id": 5, "Name": "Singkawang"}, {"Id": 6, "Name": "Bau-Bau"}, {"Id": 6, "Name": "Gorontalo"}, {"Id": 6, "Name": "Kendari"}, {"Id": 6, "Name": "Makasar"}, {"Id": 6, "Name": "MAMUJU"}, {"Id": 6, "Name": "Manado"}, {"Id": 6, "Name": "Palopo"}, {"Id": 6, "Name": "Palu"}, {"Id": 7, "Name": "Banda Aceh"}, {"Id": 7, "Name": "Bandar Lampung"}, {"Id": 7, "Name": "Batam"}, {"Id": 7, "Name": "Baturaja"}, {"Id": 7, "Name": "Bengkalis"}, {"Id": 7, "Name": "Bengkulu"}, {"Id": 7, "Name": "Binjai"}, {"Id": 7, "Name": "Jambi"}, {"Id": 7, "Name": "LAHAT"}, {"Id": 7, "Name": "Lubuklinggau"}, {"Id": 7, "Name": "Medan"}, {"Id": 7, "Name": "Padang"}, {"Id": 7, "Name": "Palembang"}, {"Id": 7, "Name": "Pekanbaru"}, {"Id": 7, "Name": "Prabumulih"}, {"Id": 7, "Name": "Riau"}, {"Id": 7, "Name": "Tanjung Pinang"}, {"Id": 8, "Name": "Badung"}, {"Id": 8, "Name": "Denpasar"}, {"Id": 8, "Name": "Mataram"}, {"Id": 9, "Name": "Ambon"}, {"Id": 9, "Name": "Jayapura"}, {"Id": 9, "Name": "Kupang"}]


url = "http://www.matahari.co.id/store-locator-ajax"

start = 0
end = 0

for y in range(0, len(listCities)+1, 5):
    start = y
    end = start + 5




    for x in range(start, end + 1):
        payload = {
            '_token': '9KzVu795HHsAiVTF00lgNYMEnqJnxv3OhwEUmRWP',
            'type': 'city',
            'city': listCities[x]['Name'],
            'id': str(listCities[x]['Id'])
        }

        res = requests.request("POST", url, headers=headers, data=payload)

        if res.status_code == 200:

            output = res.json()

            soup = BeautifulSoup(output['storeLocationView'], "html.parser")

            try:

                outputLeft = soup.find('div', class_='mCustomScrollbar').find_all('h3')
                outputRight = soup.find('div', class_='mCustomScrollbar').find_all('p')

                for i in range(0, len(outputLeft)):
                    obj = OBJ()

                    try:
                        obj.Title = outputLeft[i].text
                    except:
                        obj.Title = ""

                    try:
                        obj.Address = outputRight[i].text
                    except:
                        obj.Address = ""

                    print(obj.Title + " | " + obj.Address)

                    result = False

                    if len(listOBJ) > 0:
                        for i in range(len(listOBJ)):
                            if (str(obj.Title) == str(listOBJ[i].Title) and str(obj.Address) == str(
                                    listOBJ[i].Address)):
                                result = True
                                break

                    if result == False:
                        try:
                            y = str(obj.Title) + ' ' + str(obj.Address)

                            location = geolocator.geocode(
                                y.replace(',', ', ').translate(str.maketrans('', '', string.punctuation)))

                            obj.Latitude = str(location.latitude)
                            obj.Longitude = str(location.longitude)
                        except:
                            obj.Latitude = ''
                            obj.Longitude = ''

                        listOBJ.append(obj)

            except:
                listError.append(url)
                continue

        # break

        dt_string = now.strftime("%Y%m%d%H%M%S")
        workbook = xlsxwriter.Workbook('Documents\\Matahari_ID_' + dt_string + '.xlsx')
        worksheet = workbook.add_worksheet()

        j = 0

        for z in range(len(listOBJ)):
            j = j + 1
            print(str(j) + " of " + str(len(listOBJ)) + " | " + listOBJ[z].Title + " | " + listOBJ[
                z].Address + " | " + str(listOBJ[z].Latitude) + " | " + str(listOBJ[z].Longitude))
            worksheet.write('A' + str(j), str(listOBJ[z].Title))
            worksheet.write('B' + str(j), str(listOBJ[z].Address))
            worksheet.write('C' + str(j), str(listOBJ[z].Latitude))
            worksheet.write('D' + str(j), str(listOBJ[z].Longitude))

        workbook.close()








