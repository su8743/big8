import os
import pandas as pd
from pandas import Series, DataFrame
import xlwings as xw

cwd = os.getcwd()
print(cwd)

df_계정과목 = pd.read_excel('Journal.xlsx',sheet_name = "계정과목")
df_분개 = pd.read_excel('Journal.xlsx',sheet_name = "분개")
df1 = df_분개.groupby(['계정과목명'])
df2 = df1[['차변','대변']].sum()
print(df2)
df2["차변잔액"] = ""
df2["대변잔액"] = ""
print(df2)
df2 = df2.reset_index()
print(df2)
df2['차변'][1]
for i in df2.index:
    if df2['차변'][i] >=  df2['대변'][i]:
        df2['차변잔액'][i] = df2['차변'][i] - df2['대변'][i]
    else :
        df2['대변잔액'][i] = abs(df2['차변'][i] - df2['대변'][i])

df3 = df2.drop(1)
df4 = df3.drop('차변잔액', axis = 1)

print(df2[0:3])
print(df2['차변']>5000000)
print(df2[df2['차변']>5000000])
print(df_분개[df_분개['계정과목명'] == '현금'])


print(df_분개['차변'].apply(lambda x: x + 200 if x > 1000000 else x))
print(df_분개.sort_values(ascending=[True,True], by=["계정과목명", "차변"]))
print(df_분개.rank(ascending=True, method = 'min'))
print(df_분개['차변'].diff())



wb = xw.Book()
sht1 = wb.sheets['Sheet1']
sht1.range('A1').value = df2
