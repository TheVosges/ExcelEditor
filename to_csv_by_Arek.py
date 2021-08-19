# -*- coding: utf-8 -*-
"""
Created on Mon Jun 28 10:18:43 2021

@author: R252202
"""
import pandas as pd
from openpyxl.workbook import Workbook
from datetime import datetime

df = pd.read_clipboard()
#print(df)

wanted_values = df[["Product Value", "Product Description", "Value", "UOM", "Start Date", "End Date"]]


today = datetime.date(datetime.now())
print (today)
i=0
indexes = []
for cell in wanted_values["End Date"]:
    i+=1
    try:
        #day = cell[:2]
        #month = cell[3:6]
        #datetime_object = datetime.strptime(month, "%b")
        #month_no = datetime_object.month
        #year = "20" + str(cell[7:])
        #date_str = str(year) + "-" + str(month_no) + "-" + str(day)
        if str(cell) != "nan":
            cell = str(cell)
            month_start = cell.find('-')
            month = cell[month_start+1:month_start+4]
            datetime_object = datetime.strptime(month, "%b")
            month_no = datetime_object.month
            date_obj = str(cell[:month_start]) + "-" + str(month_no) + "-" + "20" + str(cell[month_start+5:])
            date_obj = datetime.strptime(date_obj, "%d-%m-%Y")
            date_obj = date_obj.date()
            if date_obj < today:
                indexes.append(i-1)
                
        else:
    
            continue
    except TypeError:
        continue

wanted_values = wanted_values.drop(indexes)

stored = wanted_values.to_excel('Please_rename.xlsx', index = None )