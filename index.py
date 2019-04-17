import xlrd
import os
import requests
import re
from xlutils.copy import copy  # http://pypi.python.org/pypi/xlutils
from xlwt import easyxf
import pandas as pd
from bs4 import BeautifulSoup
import datetime
import balloontip

loc = ("pincode.xlsx")
url = "https://www.mapsofindia.com/pincode/"
alternateUrl = "http://pin-code.in/advance_search.jsp?qry="

output_file_path = "D:/python/pincode_lookup/out/"

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)  # read only copy
wb = copy(wb) # a writable copy
w_sheet = wb.get_sheet(0)  # the sheet to write to within the writable copy
sheet.cell_value(0, 1)

def lookup_fn(df, key_row, key_col):
    try:
        return df.iloc[key_row, key_col]
    except IndexError:
        return 0

startTime = datetime.datetime.now()
print("\n============== Started: " + str(startTime) + "==============\n")

for i in range(sheet.nrows):
    val = (sheet.cell_value(i, 0))
    if val is '':
        break
    else:
        pincode = str(val)[:6] 
        print(str(i) + " : " + url + pincode)
        headers = {"Content-Type": "application/json"}
        response = requests.get(url + pincode + "/", headers=headers)
        respText = str(response.content)
        tables = pd.read_html(respText)
        sp500_table = tables[0]
        df = pd.DataFrame(sp500_table)
        state = lookup_fn(df, 1, 0)  # df.iloc[1,0]
        city = lookup_fn(df, 1, 1)  # df.iloc[1,1]
        if state == 0:
            response = requests.get(alternateUrl + pincode, headers=headers)
            respText = str(response.content)
            bs = BeautifulSoup(respText,"lxml")
            rows = bs.find_all("table", class_="view_codes")
            for tag in rows:
                trTags = tag.find_all("a")
                print(str(i) + " : " + alternateUrl + pincode)
                index = 0
                for tag in trTags:
                    if index == 0:
                        state = tag.text
                    if index == 2:
                        city = tag.text
                    index += 1

        w_sheet.write(i, 7, state)
        w_sheet.write(i, 8, city)

endTime = datetime.datetime.now()
wb.save(output_file_path + "data.csv" + os.path.splitext(output_file_path)[-1])
print("\n================== Time Taken: "  + str(endTime - startTime) + "===================")
balloontip.balloon_tip("Pincode lookup","Finished lookup in " + str(endTime))
