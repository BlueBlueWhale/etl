import pandas as pd

ws_contract_info = pd.read_excel('中粮屯河.xlsx', sheet_name='合同基础数据')
ws_farmer_info = pd.read_excel('contract21-22榨季银联保组汇总第一至十八批涉及蔗农.xlsx', sheet_name='涉及蔗农')
df  = pd.DataFrame(columns = ws_contract_info.columns) # pandas.DataFrame.columns The column labels of the DataFrame.

for i in range(ws_contract_info.shape[0]):
    find = False
    for j in range(ws_farmer_info.shape[0]):
        if ws_contract_info.iat[i,3] == ws_farmer_info.iat[j,0]:
            find = True
            append = ws_contract_info.iloc[i]
            append[5] = ws_farmer_info.iat[j,2]
            append[6] = ws_farmer_info.iat[j,1]
            append[7] = ws_farmer_info.iat[j,4]
            append[8] = ws_farmer_info.iat[j,8]
            append[9] = ws_farmer_info.iat[j,7]
            df = df.append(append)
    if find == False:
        df = df.append(ws_contract_info.iloc[i])
df.to_excel("合同基础数据.xlsx")
#with ExcelWriter("合同基础数据.xlsx") as writer:
#    df.to_excel(writer, sheet_name="Sheet1")

#from openpyxl import load_workbook
#from openpyxl import Workbook

#wb = load_workbook('合浦湘桂.xlsx')
#ws = wb["蔗农信息"]
#ws_to_list = list(ws.values)

 
##Find repeated data
#len = len(ws_to_list)
#for i in range(1,len):
#    for j in range(i+1,len):
#        if ws_to_list[i][5] == ws_to_list[j][5]:
#            print('重复：'+ ws_to_list[i][5])
