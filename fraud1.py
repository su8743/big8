# -*- coding: utf-8 -*-
"""
Created on Mon May  7 11:58:11 2018

@author: 김태식
"""
import os
import pandas as pd

cwd = os.getcwd()

df1 = pd.read_excel('JE2016.xlsx')
df1.sort_values(['JDate','JNo'], ascending = [False,False])
print(df1.head())
df_stratify = df1.fillna(0).groupby(['ClientCode'])

for i in df_stratify['DrAmount']:
    print(df_stratify['DrAmount'].sum().apply("{:,}".format))


    
    