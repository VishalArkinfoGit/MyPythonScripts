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
from openpyxl.descriptors.base import Integer
geolocator = Nominatim(user_agent="MyGeoCoder")
import xlsxwriter
from datetime import datetime
import string

now = datetime.now()



class OBJ:
    Id = ""
    Suburb = ""
    Postcode = ""
    PostCodeName = ""
    Title = ""
    Address = ""
    FullAddress = ""
    # Suburb = ""
    State = ""
    City = ""
    Country = ""
    # Postcode = ""
    Latitude = ""
    Longitude = ""

listOBJ = []


dt_string = now.strftime("%d%m%Y")

payload={}
headers = {
  'cookie': '__utmz=18222588.1618502314.1.1.utmcsr=google|utmccn=(organic)|utmcmd=organic|utmctr=(not provided); _fbp=fb.2.1618502315516.386669049; _gcl_au=1.1.716371224.1618502317; _ga=GA1.1.1587429724.1618502318; _hjid=f6d86cab-91cc-456f-bc95-aad8baa38fd8; __tawkuuid=e::trailerrentals.com.au::PYksjnoALhSLtwKXRt7KMbFIyNd4rNdwrPvfXJVXlA25RVnfewmnS/0MyQf40g8X::2; __utma=18222588.1558827194.1618502314.1618506614.1618510241.3; _ga_W21MCKND2E=GS1.1.1618510239.3.1.1618510241.0; eRent_msj=NSdhmpm2XQgWS9EeETGoHUXNpTacUHdf0O6oU/ZfKkE=',
  'dnt': '1',
  'referer': 'https://www.trailerrentals.com.au/location/near-by?suburb=ABECKETT+STREET'
}


listAustraliaLocations = []

for x in range(10, 6, -1):
    url = 'https://www.trailerrentals.com.au/Booking/GetPostCodeLocations?filter%5Bfilters%5D%5B0%5D%5Bfield%5D=PostCode&location='+str(x)

    print(url)

    res = requests.request("GET", url, headers=headers, data=payload)

    if res.status_code == 200:
        output = json.loads(res.text)

        print(len(output))

        for y in output:
            obj = OBJ()

            obj.Id = y['Id']
            obj.Postcode = y['PostCode']
            obj.PostCodeName = y['PostCodeName']
            obj.Suburb = y['Suburb']

            result = False

            if len(listAustraliaLocations) > 0:
                for z in range(len(listAustraliaLocations)):
                    if (str(obj.Id) == str(listAustraliaLocations[z].Id) and str(obj.Postcode) == str(listAustraliaLocations[z].Postcode)):
                        result = True
                        break

            if result == False:
                listAustraliaLocations.append(obj)
    #     break
    # break


print(len(listAustraliaLocations))


workbook = xlsxwriter.Workbook('Documents\\UHAUL_Cities_AUS_' + dt_string + '.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Title')
worksheet.write('B1', 'Address')
worksheet.write('C1', 'City')
worksheet.write('D1', 'State')

j = 1
for z in range(len(listAustraliaLocations)):
    j = j + 1
    print(str(listAustraliaLocations[z].Id) + " | " + str(listAustraliaLocations[z].Postcode) + " | " + str(listAustraliaLocations[z].PostCodeName))

    worksheet.write('A' + str(j), str(listAustraliaLocations[z].Id))
    worksheet.write('B' + str(j), str(listAustraliaLocations[z].Postcode))
    worksheet.write('C' + str(j), str(listAustraliaLocations[z].PostCodeName))
    worksheet.write('D' + str(j), str(listAustraliaLocations[z].Suburb))

workbook.close()


# for x in listAustraliaLocations:
#     try:
#         url = 'https://www.trailerrentals.com.au/location/near-by?postcode='+str(x.PostCode)+'&locations='+str(x.PostCodeName).replace(' ', '+')+'&suburb='+str(x.Suburb).replace(' ', '+')
#
#         print(url)
#
#         res = requests.request("GET", url, headers=headers, data=payload)
#
#         if res.status_code == 200:
#             soup = BeautifulSoup(res.content, "html.parser")
#
#             listH4 = soup.find(class_="box-select-address").find_all('h4')
#             listSpan = soup.find(class_="box-select-address").find_all('span')
#
#             for i in range(0, len(listH4)):
#                 try:
#                     obj = OBJ()
#
#                     obj.Title = str(listH4[i].text)
#                     obj.Address = ''
#                     obj.FullAddress = str(listSpan[i].text)
#                     obj.City = ''
#                     obj.Country = ''
#                     obj.Postcode = ''
#                     obj.Suburb = ''
#                     obj.State = ''
#                     obj.Latitude = ''
#                     obj.Longitude = ''
#
#                     try:
#                         xy = str(obj.FullAddress).split(", ")
#                         obj.Postcode = str(xy[len(xy) - 1])
#                         obj.State = str(xy[len(xy) - 2])
#                         obj.City = str(xy[len(xy) - 3])
#
#                         zx = xy[0: (int(len(xy)) - 3)]
#                         obj.Address = ','.join(zx)
#                     except:
#                         try:
#                             xy = str(obj.FullAddress).split(", ")
#                             obj.Postcode = str(xy[len(xy) - 1])
#                             obj.State = str(xy[len(xy) - 2])
#                             # obj.City = str(y[len(y) - 3])
#
#                             zx = xy[0: (int(len(xy)) - 2)]
#                             obj.Address = ','.join(zx)
#                         except:
#                             obj.FullAddress = obj.FullAddress
#
#
#                     try:
#                         location = geolocator.geocode(obj.FullAddress.translate(str.maketrans('', '', string.punctuation)) + ' Australia')
#                         zy = str(location.address).split(", ")
#
#                         # if len(zy) >= 5:
#                         #     obj.Country = str(zy[len(zy) - 1])
#                         #     obj.State = str(zy[len(zy) - 2])
#                         #     obj.City = str(zy[len(zy) - 3])
#                         #     obj.Suburb = str(zy[len(zy) - 4])
#                         #
#                         #     zx = zy[0: (int(len(zy)) - 4)]
#                         #     obj.Address = ','.join(zx)
#                         # elif len(zy) == 4:
#                         #     obj.Country = str(zy[len(zy) - 1])
#                         #     obj.State = str(zy[len(zy) - 2])
#                         #     obj.City = str(zy[len(zy) - 3])
#                         #
#                         #     zx = zy[0: (int(len(zy)) - 3)]
#                         #     obj.Address = ','.join(zx)
#                         # else:
#                         #     obj.Postcode = obj.Postcode
#
#                         obj.Latitude = str(location.latitude)
#                         obj.Longitude = str(location.longitude)
#
#                     except:
#                         obj.FullAddress = obj.FullAddress
#
#                     print(obj.Title + " | " + obj.FullAddress + " | " + str(obj.Latitude) + " | " + str(obj.Longitude))
#
#                     result = False
#
#                     if len(listOBJ) > 0:
#                         for i in range(len(listOBJ)):
#                             # print(str(obj.Title) + " " + str(listOBJ[z].Title))
#                             if (str(obj.Latitude) == str(listOBJ[i].Latitude) and str(obj.Longitude) == str(
#                                 listOBJ[i].Longitude)):
#                                 result = True
#                                 break
#
#                     if result == False:
#                         listOBJ.append(obj)
#
#                 except:
#                     print("*******************************************")
#                     print("Location Not Found: " + str(listH4[i].text) + " | " + str(listSpan[i].text))
#                     print("*******************************************")
#                     print("\n")
#                     continue
#     except:
#         print("*******************************************")
#         print("Location Not Found: " + url)
#         print("*******************************************")
#         print("\n")
#         continue
#
#
# workbook = xlsxwriter.Workbook('Documents\\UHAUL_AUS_' + dt_string + '.xlsx')
# worksheet = workbook.add_worksheet()
#
# worksheet.write('A1', 'Title')
# worksheet.write('B1', 'Address')
# worksheet.write('C1', 'City')
# worksheet.write('D1', 'State')
# worksheet.write('E1', 'Postcode')
# worksheet.write('F1', 'Country')
# worksheet.write('G1', 'Latitude')
# worksheet.write('H1', 'Longitude')
# worksheet.write('I1', 'FullAddress')
#
# j = 1
# for z in range(len(listOBJ)):
#     j = j + 1
#     print(listOBJ[z].Title + " | " + listOBJ[z].Address + " | " + listOBJ[z].City + " | " + listOBJ[
#         z].State + " | " + str(listOBJ[z].Postcode) + " | " + listOBJ[z].Country + " | " + str(
#         listOBJ[z].Latitude) + " | " + str(listOBJ[z].Longitude))
#
#     worksheet.write('A' + str(j), str(listOBJ[z].Title))
#     worksheet.write('B' + str(j), str(listOBJ[z].Address))
#     worksheet.write('C' + str(j), str(listOBJ[z].City))
#     worksheet.write('D' + str(j), str(listOBJ[z].State))
#     worksheet.write('E' + str(j), str(listOBJ[z].Postcode))
#     worksheet.write('F' + str(j), str(listOBJ[z].Country))
#     worksheet.write('G' + str(j), str(listOBJ[z].Latitude))
#     worksheet.write('H' + str(j), str(listOBJ[z].Longitude))
#     worksheet.write('I' + str(j), str(listOBJ[z].FullAddress))
#
# workbook.close()
