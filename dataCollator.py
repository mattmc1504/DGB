# program that takes rough data given from proprietary software and collates it under relevant groups
# printed and filtered, ready to sort and compare

import openpyxl 
path = "test.xlsx"
wb = openpyxl.load_workbook(path) 
sheet = wb.active 
sheet2 = wb.create_sheet (title = "Collated Data")

row = sheet.max_row 
column = sheet.max_column 

for i in range(1, row + 1): 
    item = sheet.cell(row=i, column=2)
    if item.value == None:
        continue 
    net_sales = sheet.cell(row=i, column=3).value
    percent = sheet.cell(row=i, column=4).value     
    item_following = sheet.cell(row=i+1, column=2).value
    if item_following == None:
        for x in range(1, row-i+1):
            item_iteration = sheet.cell(row=i+x, column=2)
            if item_iteration.value != None:
                break
            if sheet.cell(row=i+x, column=3).value == None:
                percent += sheet.cell(row=i+x, column=4).value
                continue    
            elif sheet.cell(row=i+x, column=4).value == None:     
                net_sales += sheet.cell(row=i+x, column=3).value
                continue                    
            else:
                net_sales += sheet.cell(row=i+x, column=3).value                
                percent += sheet.cell(row=i+x, column=4).value     
    c1 = sheet2.cell(row = i, column = 2)
    c1.value = str(item.value)
    c2 = sheet2.cell(row = i, column = 3)
    c2.value = (net_sales)
    c3 = sheet2.cell(row = i, column = 4)
    c3.value = (percent)
    
filters = sheet2.auto_filter
filters.ref = "B2:D1000"
