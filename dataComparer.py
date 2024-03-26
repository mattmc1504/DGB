# program that takes collated data and produces percentage change figures
# based on another set of data

import openpyxl 
path = "test.xlsx"
wb = openpyxl.load_workbook(path) 
sheet = wb.active 
sheet2 = wb ["Collated Data"]

row = sheet.max_row 
column = sheet.max_column 

for i in range(1, row + 1): 

    item1 = sheet.cell(row=i, column=2).value
    net_sales1 = sheet.cell(row=i, column=3).value
    percent1 = sheet.cell(row=i, column=4).value         
    for x in range(1, row + 1):
        item2 = sheet.cell(row=x, column=6).value
        net_sales2 = sheet.cell(row=x, column=7).value
        percent2 = sheet.cell(row=x, column=8).value    
        if item2 == None:
            continue
        if item1 == item2:
            sales_change = (net_sales1 / net_sales2)
            percent_change = (percent1 / percent2)
            c1 = sheet2.cell(row = i, column = 9)
            c1.value = (sales_change)
            c2 = sheet2.cell(row = i, column = 10)
            c2.value = (percent_change)

wb.save (path)