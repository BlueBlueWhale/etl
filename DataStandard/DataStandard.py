# -*- coding: utf-8 -*-
"""
Created on Sat Sep 18 17:41:15 2021

@author: wei
"""
# %% Initialization.
from data import excel as e

wb = e.Worksheets('数据标准（订单，蔗农档案，品种，种植情况等）2021.9.xls')
# %% Check reused abbreviations.
for i in range(wb.count):
    mismatch_count = 0
    for j in range(i+1, wb.count):
        for x in range(3, wb.row_count(i)):
            for y in range(3, wb.row_count(j)):
                col = range(1,3)
                bools = wb.getdata(i, x, col) == wb.getdata(j, y, col)
                # bools is a pandas.Series object.
                if bools[0]^bools[1]:   # ^ is XOR operator.
                    mismatch_count += 1
                    if mismatch_count == 1:
                        print("sheet ", i, ", row ", x, "\n", wb.getdata(i, x, col), "\n")
                    print("sheet ", j, ", row ", y, "\n", wb.getdata(j, y, col), "\n")
                    print("^^^^^^", mismatch_count, "mismatch", "^^^^^^", "\n")
