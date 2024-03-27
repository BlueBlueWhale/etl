
from openpyxl import load_workbook
from openpyxl import Workbook

wb = load_workbook('中粮屯河.xlsx')
ws = wb["合同基础数据"]
ws_to_list = list(ws.values)

wb_rep = Workbook()
ws_rep = wb_rep.active
ws_rep.title = "合同基础数据"
ws_rep.append(ws_to_list[0])

 
#Find repeated data
len = len(ws_to_list)
for i in range(1,len):
    for j in range(i+1,len):
        if ws_to_list[i][5] == ws_to_list[j][5]:
            print('重复：'+ ws_to_list[i][5])
            ws_rep.append(ws_to_list[i])

wb_rep.save('中粮屯河重复数据.xlsx')