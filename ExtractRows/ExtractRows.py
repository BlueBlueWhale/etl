# -*- coding: utf-8 -*-
"""
Created on Wed Dec 22 21:48:50 2021

@author: wyz
"""

# %% Initialization.
import sys
sys.path.append(r'D:/working-directories/python/packages')
from data import excel as e

wb = e.Workbook('设备清单.xlsx')

# %% Replace TABLE_NAME.
for i in wb.count:
    range(ziduan.row_count(0)):
    for j in range(biaoming.row_count(0)):
        if ziduan.getdata(0, i, 0) == biaoming.getdata(0, j, 0):
            ziduan.putdata(0, i, 0, biaoming.getdata(0, j, 2))

# ziduan.save_sheets('ziduan.xlsx')