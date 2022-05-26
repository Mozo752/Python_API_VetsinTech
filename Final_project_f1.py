import requests
import json
import openpyxl
from openpyxl import load_workbook

# 
# assign url to variable
url = "https://ergast.com/api/f1/current/last/laps.json"

response = requests.request("GET", url)

#check response status of url
#print response 
# print(response)

#print content / response.text
# print(response.text)

# #Assign response object to variable (json loads)
clean_data = json.loads(response.text)
# # #print the cleaned data
# print(clean_data)
# print(clean_data)

clean1 = clean_data['MRData']
# # print(clean1)

clean2 = clean1['RaceTable']
# print(clean2)

clean3 = clean2['Races']
print(clean3)

# def getList(dict):
#     return dict.keys()

# dict = clean3
# print(getList(dict))
 

# for i in range(len(clean3)):
#     result = clean3[i]['driverId']
#     print(result)
    
# #create a workbook
# wb = load_workbook('empty_book.xlsx')
# ws = wb.active
# #Create a spreadsheet page
# sheet = wb.create_sheet("sheet1")


# #Designate columns
# sheet['A1'] = "id"
# # sheet['A1'].font = openpyxl.styles.Font(bold=True)
# sheet['B1'] = "name"
# ws.title = "Motorcylces"

# header_list = [sheet['A1'], sheet['B1']]

# #Loop through the columns
# for i in range(len(header_list)):
#     #change the font to bold
#     header_list[i].font = openpyxl.styles.Font(bold=True)
    
# #Loop through data from api
#     #fill cell falues with appropriate data
    
# #Save the workbook
# wb.save('empty_workbook.xlsx')