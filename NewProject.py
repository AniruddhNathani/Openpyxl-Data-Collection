# -*- coding: utf-8 -*-
"""
Created on Wed Jun 20 14:08:44 2018

@author: nathani_n
"""

import pandas as pd
from openpyxl import Workbook
from openpyxl import load_workbook

from openpyxl.utils import coordinate_from_string, column_index_from_string


print("\n\nEnter The number of Data Items you want to search:")
num=int(input())
print("\n\nEnter Items you want to Get Data for:")

list_search_items=['ModuleId']
for _ in range(num): 
    list_search_items.append(input())




wb1=Workbook()
dest_filename=r'C:\Users\nathani_n\Desktop\ExcelData\FinalDisplay.xlsx'
ws1 = wb1.active
ws1.title = "Final Data"
ws1.append(list_search_items)


wb2= Workbook()                
wb2= load_workbook(r'C:\Users\nathani_n\Desktop\ExcelData\Predictive_Analytics_Consolidated_MIM_Information_v1.xlsx')
ws2=wb2[wb2.sheetnames[0]]
xy=[]
row_main=[]
for i in range(1,num+1):
    for row_i in range(1, ws2.max_row + 1):
        for column_i in range(1,3):
            if ws2.cell(row=row_i, column=column_i).value == list_search_items[i]:
                row_main.append(row_i)
    
list_new=[]
cell_values_sheet=[]
for sheets in wb2:
    cell_values=[]
    cell_values.append(str(sheets))
    for i in range(len(row_main)):
        for row in sheets.iter_rows('B{}'.format(row_main[i])):
                for col in sheets.iter_cols(min_col=2,max_col=2):
                    for cell in row:
                        cell_values.append(cell.value)
    ws1.append(cell_values)      
    wb1.save(filename = dest_filename)






