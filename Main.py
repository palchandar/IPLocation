# -*- coding: utf-8 -*-
"""
Created on Mon Feb 19 10:54:30 2018

@author: palchandar Subbarao

"""
import IP2Location
import xlrd
import csv

filename ="SampleValidIPAddresses.xlsx"
''' Give valid Ipaddress in the excel format '''

sheetname = "ipsample"
''' Provide sheetname in the current excel file '''

workbook = xlrd.open_workbook(filename)
worksheet = workbook.sheet_by_name(sheetname)

num_rows = worksheet.nrows
num_cols = worksheet.ncols

result_res =[]
for curr_row in range(1, num_rows, 1):
    row_res = []
    
    for curr_col in range(0, num_cols, 1):
       data = worksheet.cell_value(curr_row, curr_col)
       database = IP2Location.IP2Location("IP2LOCATION-LITE-DB1.BIN")
       rec = database.get_all(data)
       country = rec.country_long
       result = data + "==>" + country
       row_res.append(result)
       
    result_res.append(result)

#print result_res
f = open('data.csv', 'wb')
''' Provide outputfile name here '''
w = csv.writer(f, delimiter = ',')
w.writerow(["IPAddress","Country"])
w.writerows([x.split('==>') for x in result_res])
f.close()