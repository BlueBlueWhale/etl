# -*- coding: utf-8 -*-
"""
Created on Tue Dec  7 15:10:09 2021

@author: wyz
"""
# %% Initialization.
import pandas as pd
import re

df = pd.read_excel(r'来宾2020良补验收表数据.xlsx', skiprows=(1))
# %% Data cleaning: delete whitespaces.
print('去掉', df['蔗农姓名'].str.contains(' ').sum(), '条数据\'蔗农姓名\'字段首尾的空格。')
df['蔗农姓名'] = df['蔗农姓名'].str.strip()
# %% Check for empty names.
empty_name = df[df['蔗农姓名'] =='']
# %% Check for invalid ids.
# invalid_id = df[list(map(lambda x: len(x) != 18,df['身份证']))]
invalid_id = df[df['身份证'].apply(
    lambda x: len(str(x)) != 18 or bool(re.search(r'(.)\1{17}', x))
    )]
# %% Check for one-to-many mappings from '身份证' to '蔗农姓名'
repeated_id = pd.DataFrame()
group_by_id = df.groupby(['身份证'])
for name, group in group_by_id:
    # group_by_id.get_group('450521199312022510')['蔗农姓名'].unique()
    if group['蔗农姓名'].nunique() != 1:
        repeated_id = repeated_id.append(group)
# %% Check for one-to-many mappings from '蔗农姓名' to  '身份证'
repeated_name = pd.DataFrame()
group_by_name = df.groupby(['蔗农姓名'])
for name, group in group_by_name:
    if group['身份证'].nunique() != 1:
        repeated_name = repeated_name.append(group)
# %% Add comments.
df['备注'] = ''
df.loc[invalid_id.index,'备注'] = '身份证号错误，'
df.loc[repeated_id.index,'备注'] = df.loc[repeated_id.index,'备注'] + '身份证号对应多个姓名，'
df.loc[repeated_name.index,'备注'] = df.loc[repeated_name.index,'备注'] + '姓名对应多个身份证号，'
# %% save to excel
df.to_excel(r'来宾2020良补验收表数据备注问题.xlsx', index = False)
# =============================================================================
# with pd.ExcelWriter('来宾2021订单合同问题数据.xlsx', mode='w') as writer:
#     empty_name.to_excel(writer, sheet_name = '姓名为空')
#     invalid_id.to_excel(writer, sheet_name = '身份证号错误')
#     repeated_id.to_excel(writer, sheet_name = '同身份证号不同姓名')
#     repeated_name.to_excel(writer, sheet_name = '同姓名不同身份证号')
# =============================================================================
