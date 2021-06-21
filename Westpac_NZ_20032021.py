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

now = datetime.now()

# dt_string = now.strftime("%Y%m%d%H%M%S")
dt_string = now.strftime("%Y%m%d")

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

listOBJ = []

listNotFound = []
listAll = []

j = 0

for i in range(0, 555, 12):

    url = 'https://www.westpac.co.nz/contact-us/branch-finder?start=' + str(i)
    res = requests.get(url)
    # print(url)
    if res.status_code == 200:
        try:
            soup = BeautifulSoup(res.content, "html.parser")
            output = soup.find(class_="finder-listing").find_all(class_="js-card")
            # print(output)

            listAddress = []

            for z in output:
                y = z['data-props']

                y = json.loads(y)

                listAddress.append(y)

            output = listAddress

            # print(output)

            for x in output:

                obj = OBJ()

                try:
                    obj.Title = x['siteName']
                except:
                    obj.Title = ""

                try:
                    obj.Address = x['address']
                except:
                    obj.Address = ""
                try:
                    obj.FullAddress = x['locationType']
                except:
                    obj.FullAddress = ""

                try:
                    obj.Latitude = x['latitude']
                except:
                    obj.Latitude = ""
                try:
                    obj.Longitude = x['longitude']
                except:
                    obj.Longitude = ""

                try:
                    search = str(obj.Latitude) + ", " + str(obj.Longitude)

                    location = geolocator.reverse(search)
                    y = str(location.address).split(", ")
                    if len(y) >= 5:
                        obj.Postcode = str(y[len(y) - 2])
                        obj.State = str(y[len(y) - 3])
                        obj.City = str(y[len(y) - 4])
                        obj.Suburb = str(y[len(y) - 5])
                    elif len(y) == 4:
                        obj.Postcode = str(y[len(y) - 2])
                        obj.State = str(y[len(y) - 3])
                        obj.City = str(y[len(y) - 4])
                    else:
                        obj.Postcode = ""

                except:
                    obj.Postcode = ""

                obj.Country = "New Zeland"

                print(obj.Title + " | " + obj.Suburb + " | " + obj.State + " | " + str(
                    obj.Postcode) + " | " + obj.Country + " | " + str(obj.Latitude) + " | " + str(obj.Longitude))
                # print(obj.Title)


                result = False

                if len(listOBJ) > 0:
                    for z in range(len(listOBJ)):
                        if (str(obj.Latitude) == str(listOBJ[z].Latitude) and str(obj.Longitude) == str(listOBJ[z].Longitude)):
                            result = True
                            break

                if result == False:
                    listOBJ.append(obj)

            #     # break
        except:
            j=j+1
            print("*******************************************")
            print("Location Not Found: " + url)
            listNotFound.append(url)
            print("*******************************************")
            print("\n")

    # if i == 1:
    #     break

print("Total Address: "+ str(len(listOBJ)))


workbook = xlsxwriter.Workbook('Documents\\Westpac_NZ_'+dt_string+'.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write('A1', 'Title')
worksheet.write('B1', 'Address')
worksheet.write('C1', 'Address2')
worksheet.write('D1', 'Suburb')
worksheet.write('E1', 'City')
worksheet.write('F1', 'State')
worksheet.write('G1', 'Postcode')
worksheet.write('H1', 'Country')
worksheet.write('I1', 'Latitude')
worksheet.write('J1', 'Longitude')

j = 1
for z in range(len(listOBJ)):
    j = j + 1
    print(listOBJ[z].Title + " | " + listOBJ[z].Address + " | " + listOBJ[z].City + " | " + listOBJ[
        z].State + " | " + str(listOBJ[z].Postcode) + " | " + listOBJ[z].Country + " | " + str(
        listOBJ[z].Latitude) + " | " + str(listOBJ[z].Longitude))

    worksheet.write('A' + str(j), str(listOBJ[z].Title))
    worksheet.write('B' + str(j), str(listOBJ[z].Address))
    worksheet.write('C' + str(j), str(listOBJ[z].FullAddress))
    worksheet.write('D' + str(j), str(listOBJ[z].Suburb))
    worksheet.write('E' + str(j), str(listOBJ[z].City))
    worksheet.write('F' + str(j), str(listOBJ[z].State))
    worksheet.write('G' + str(j), str(listOBJ[z].Postcode))
    worksheet.write('H' + str(j), str(listOBJ[z].Country))
    worksheet.write('I' + str(j), str(listOBJ[z].Latitude))
    worksheet.write('J' + str(j), str(listOBJ[z].Longitude))

workbook.close()

print(listNotFound)
