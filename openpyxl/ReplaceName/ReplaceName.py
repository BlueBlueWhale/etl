from openpyxl import load_workbook
from openpyxl import Workbook

wb_data = load_workbook('数据模版.xlsx')
wb_dic = load_workbook('数据字典.xlsx')
ws_dic = wb_dic.active

keys = ws_dic['C']
values = ws_dic['D']
#values = list(ws_dic.iter_cols(4,4))[0]
for key in keys:
    if key.value == None:
        print("Empty row: "+str(key.row))
        continue
    if key.value.find('[') != -1:
        key.value = key.value[1:-1]
for key in keys:
    print(key.value)

for worksheet in wb_data.worksheets:
    column_headings = worksheet[1]
    for column_heading in column_headings:
        for i in range(len(keys)):
            if column_heading.value == keys[i].value:
                column_heading.value = values[i].value
                #print("renamed")
                break
wb_data.save('数据.xlsx')

#wb2.active.iter_cols(1)
#cell(row=, volumn=3)
#j.value
#    list(wb1.worksheets[0].columns)

#wb1.worksheets[i].cell(row=1, column=).value
#Python’s built-in function chr () is used for converting an Integer to a Character, while the function ord () is used to do the reverse, i.e, convert a Character to an Integer.