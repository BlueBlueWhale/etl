# -*- coding: utf-8 -*-
"""
Created on Mon Dec  6 11:32:25 2021

@author: wyz
"""
# %% Initialization.
import sys
sys.path.append(r'D:/working-directories/python/packages')
from data import excel as e

ziduan = e.Workbook('字典，字段.xlsx')
biaoming = e.Workbook('字典，表名.xlsx')

# %% Replace TABLE_NAME.
for i in range(ziduan.row_count(0)):
    for j in range(biaoming.row_count(0)):
        if ziduan.getdata(0, i, 0) == biaoming.getdata(0, j, 0):
            ziduan.putdata(0, i, 0, biaoming.getdata(0, j, 2))

# ziduan.save_sheets('ziduan.xlsx')

# %% Reformat table.
ziduan1 = e.Workbook('ziduan1.xlsx')
table = ""
row = 1
for i in range(ziduan1.column_count(0)):
    if table != ziduan1.getdata(0, 0, i):
        row = row + 1
        table = ziduan1.getdata(0, 0, i)
        ziduan1.putdata(0, row, 0, table)
    else:
        
    
ziduan1.save_sheets('ziduan2.xlsx')