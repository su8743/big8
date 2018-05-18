# -*- coding: utf-8 -*-
"""
Created on Mon May  7 11:58:11 2018

@author: 김태식
"""
import os
cwd = os.getcwd()
import pandas as pd
import xlwings as xw
import calendar


df1 = pd.read_excel('JE2016.xlsx')

wb = xw.Book()
sht1 = wb.sheets['Sheet1']
sht1.range('A1').value = df1


df2 = df1.groupby(['day'])
df3 = df2['DrAmount','CrAmount'].count()
sht2 = wb.sheets['Sheet2']
sht2.range('A1').value = df3




df1.sort_values(['JDate','JNo'], ascending = [False,False])
print(df1.head())
df_stratify = df1.fillna(0).groupby(['ClientCode'])

for i in df_stratify['DrAmount']:
    df2 = df_stratify['DrAmount'].sum().apply("{:,}".format)


    
    