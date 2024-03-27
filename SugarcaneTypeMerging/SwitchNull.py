# -*- coding: utf-8 -*-
"""
Created on Fri Oct  8 17:43:12 2021

@author: wei
"""

# %% Initialization.
import sys
sys.path.append(r'C:\Users\wei\working-directories\python\packages')
from data import excel as e

wb = e.Workbook(r'数据标准.xlsx')
# %% Switch null to not null.
for sheet in range(wb.count):
    for row in range(3, wb.row_count(sheet)):
        print(wb.getdata(sheet, row, 8))
        if wb.getdata(sheet, row, 8) == '是':
            wb.putdata(sheet, row, 8, '否')
        elif wb.getdata(sheet, row, 8) == '否':
            wb.putdata(sheet, row, 8, '是')
        print(wb.getdata(sheet, row, 8))
wb.save_sheets('数据标准.xlsx')