# -*- coding: utf-8 -*-
"""
Created on Sat Sep 18 16:40:07 2021

@author: wei
"""
# %%
import pandas as pd

Shao_Wei = r'品种表.xlsx'
Wen_Da = r'甘蔗品种.xls'
df_Shao_Wei = pd.read_excel(Shao_Wei, sheet_name = "Sheet1") # sheet_name不指定时默认返回全表数据
df_Wen_Da = pd.read_excel(Wen_Da, sheet_name = "广西甘蔗品种表")

# %%
print(df_Shao_Wei.head(), '\n\n') # 打印头部数据，仅查看数据示例时常用
print(df_Wen_Da['品种名称'])

# %%
count_mismatch = 0
for type in df_Shao_Wei['品种名称']:
    if df_Wen_Da['品种名称'].str.contains(type) is False:
        count_mismatch += 1
    # print(df_Wen_Da.loc[df_Wen_Da['品种名称'].str.contains(type),'品种名称'])
    # 
print(count_mismatch)