from pandas import Series, DataFrame
import pandas as pd
import xlwings as xw

df = pd.read_excel('JE.xlsx')

wb = xw.Book()
sht1 = wb.sheets['Sheet1']
sht1.range('A1').value = df2

#주말에 분개를 기입한(부정가능성 존재) 경우가 몇 record인지 분석

df['day'] = df['Jdate'].dt.weekday_name

df2 = df.sort_values(['ClientCode','JDate','JNo'])


df1 = df.groupby(['day'])
df1.sort_values(['JDate','JNo'], ascending = [False,False])

df_계층 = df1.fillna(0).groupby(['ClientCode'])

for i in df_계층['DrAmount']:
    df2 = df_계층['DrAmount'].sum().apply("{:,}".format)
    
