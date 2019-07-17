# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import pandas as pd

def get_col_widths(df2):
    # First we find the maximum length of the index column   
    idx_max = max([len(str(s)) for s in df2.index.values] + [len(str(df2.index.name))])
    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
    #return [idx_max] + [max([len(str(s)) for s in df2[col].values] + [len(col)]) for col in df2.columns]
    return [max([len(str(s)) for s in df2[col].values] + [len(col)]) for col in df2.columns]

df1 = pd.read_excel('C:/Users/danielyuan/Desktop/ProductPlanProject/ShipmentSchedule_04222019 - Copy.xlsx')
writer = pd.ExcelWriter('C:/Users/danielyuan/Desktop/ProductPlanProject/ShipmentSchedule_04222019-test.xlsx', engine='xlsxwriter')
df2 = df1.copy()

df2.to_excel(writer, sheet_name='Sheet1', index=False)

workbook = writer.book
worksheet = writer.sheets['Sheet1']


for i, width in enumerate(get_col_widths(df2), start=0):
    print(i, width)
    worksheet.set_column(i, i, width)

writer.save()