from openpyxl import load_workbook
from openpyxl import Workbook

#def iter_rows(ws,n):  #produce the list of items in the particular row
#        for row in ws.iter_rows(n):
#            yield [cell.value for cell in row]

# Import worksheets.
wb_land = load_workbook('地块20210810星星.xlsx')
wb_contract = load_workbook('合同导入（第二批）.xlsx')
ws_land = wb_land["Sheet2"]
ws_contract = wb_contract["Sheet1"]

# Make a new workbook and make its fields by copying from ws_land.
wb = Workbook()
ws = wb.active
values = []
for cell in ws_land[1]:
    values.append(cell.value)
ws.append(values)

# Get a set of land codes from ws_contract
contract_land_codes = set()
for cell in ws_contract['AB']:
    contract_land_codes.update( cell.value.split(';') )

for code in contract_land_codes:
    for row in ws_land:
        if code == row[1].value:
            print(code)
            values = []
            for cell in row:
                values.append(cell.value)
            ws.append(values)

wb.save('筛选出的地块数据.xlsx')

# Tample for deleting duplicates from a list
#res = []
#for i in test_list:
#    if i not in res:
#        res.append(i)