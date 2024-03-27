'''
Compare two databases and find missing data from the newer database. 
Generate an Excel file to store missing data.

To group code together, mark the code as a code cell by adding a comment starting with #%% to the beginning of the cell, which ends the previous one. Code cells can be collapsed and expanded, and using Ctrl+Enter inside a code cell sends the entire cell to the Interactive window and moves to the next one.
https://docs.microsoft.com/en-us/visualstudio/python/python-interactive-repl-in-visual-studio?view=vs-2019
'''

#%% 
from openpyxl import load_workbook
from openpyxl import Workbook

# Import the old database
wb_old = load_workbook('contract21-22榨季银联保组汇总第三批至十三批汇总.xlsx')
#print(wb_old.sheetnames)
ws_old = wb_old["合同信息"]
ws_old_to_list = list(ws_old.values)

# Import the new database
wb_new = load_workbook('中粮蔗农20210723.xlsx')
ws_new = wb_new["蔗农信息"]
ws_new_to_list = list(ws_new.values)

# Create an Excel workbook to store missing data.
wb_supp = Workbook()
ws_supp = wb_supp.active
ws_supp.title = "蔗农信息"
ws_supp.append(ws_old_to_list[0])
print(ws_supp[1])

#%% 
#Find data that are missing from the newer database, write them to a workbook, save to a file.
for i in ws_old_to_list[1:]:
    data_correct = False
    for j in ws_new_to_list[1:]:
        if i[0] == j[0]:    # if farmer names are the same
            if i[1] == j[1]:    # if farmer codes are the same
                data_correct = True
            break
    if not data_correct:
        ws_supp.append(i)

wb_supp.save('中粮屯河补充数据.xlsx')


#ws_supp.calculate_dimension()
#ws_old.dimensions
#ws_old['C11'].row
#ws_old['C11'].column
#ws_old['C11'].coordinate

#print(ws_supp.cell(row=1, column=1, value=10).value)
#len(ws_new['A'])
# type(ws_old['A'])

##ws = wb.create_sheet("MySheet0", 0)
#ws['A'][0]