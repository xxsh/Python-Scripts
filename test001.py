#!/usr/bin/env python3
import pandas as pd
from pandas import Series,DataFrame
import numpy as np
import datetime

df = pd.read_excel("SQP_updated.xlsm",sheet_name = 'Sheet2',header = 0,names=['SQ','PSQ','0','1','2','3','4','5'])
df1 = pd.DataFrame()
#df = pd.read_csv('test_input.csv',header=0,names=['SQ','PSQ','0','1','2','3','4','5'])
target_date1 = '2019-07-31'
target_date1 = datetime.datetime.strptime(target_date1,'%Y-%m-%d').date()
target_date2 = '2019-09-01'
target_date2 = datetime.datetime.strptime(target_date2,'%Y-%m-%d').date()
df1['DATE0'] = pd.to_datetime(df['0'], format="%Y-%m-%d %H:%M:%S")
df['8月数量'] = df1['DATE0'].apply(lambda x: 1 if target_date2 > x.date() > target_date1 else 0)
df1['DATE1'] = pd.to_datetime(df['1'], format= "%Y-%m-%d")
df['8月数量'] = df['8月数量'] + df1['DATE1'].apply(lambda x: 1 if target_date2 > x.date() > target_date1 else 0)
df1['DATE2'] = pd.to_datetime(df['2'], format= "%Y-%m-%d")
df['8月数量'] = df['8月数量'] + df1['DATE2'].apply(lambda x: 1 if target_date2 > x.date() > target_date1 else 0)
df1['DATE3'] = pd.to_datetime(df['3'], format= "%Y-%m-%d")
df['8月数量'] = df['8月数量'] + df1['DATE3'].apply(lambda x: 1 if target_date2 > x.date() > target_date1 else 0)
df1['DATE4'] = pd.to_datetime(df['4'], format= "%Y-%m-%d")
df['8月数量'] = df['8月数量'] + df1['DATE4'].apply(lambda x: 1 if target_date2 > x.date() > target_date1 else 0)
df1['DATE5'] = pd.to_datetime(df['5'], format= "%Y-%m-%d")
df['8月数量'] = df['8月数量'] + df1['DATE5'].apply(lambda x: 1 if target_date2 > x.date() > target_date1 else 0)
writer = pd.ExcelWriter('output.xlsx')
grouped = df['8月数量'].groupby(df['SQ']).sum()
df.to_excel(writer,'Sheet1')
grouped.to_excel(writer,'Result')
writer.save()
#print(df['8月数量'])
#print(df1)
#print(grouped)

