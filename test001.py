#!/usr/bin/env python3
import pandas as pd
import datetime


def input_judge_condition():
    starting_time = input('input starting time(eg.2019-07-31):')
    starting_time = datetime.datetime.strptime(starting_time, '%Y-%m-%d').date()
    deadline = input('input deadline (eg.2019-09-01):')
    deadline = datetime.datetime.strptime(deadline, '%Y-%m-%d').date()
    return starting_time, deadline


def calculation(starting_time, deadline):
    df = pd.read_excel("SQP_updated.xlsm", sheet_name='Sheet2', header=0,
                       names=['SQ', 'PSQ', '0', '1', '2', '3', '4', '5'])
    df1 = pd.DataFrame()
    df1['DATE0'] = pd.to_datetime(df['0'], format="%Y-%m-%d %H:%M:%S")
    df['8月数量'] = df1['DATE0'].apply(lambda x: 1 if deadline > x.date() > starting_time else 0)
    for i in ['1', '2', '3', '4', '5']:
        try:
            df1[i] = pd.to_datetime(df[i], format="%Y-%m-%d")
        except ValueError:
            df1[i] = pd.to_datetime(df[i], format="%Y-%m-%d %H:%M:%S")
            df['8月数量'] = df['8月数量'] + df1[i].apply(lambda x: 1 if deadline > x.date() > starting_time else 0)
        else:
            df['8月数量'] = df['8月数量'] + df1[i].apply(lambda x: 1 if deadline > x.date() > starting_time else 0)
    writer = pd.ExcelWriter('output.xlsx')
    df.to_excel(writer, 'Sheet1')
    grouped = df['8月数量'].groupby(df['SQ']).sum()
    grouped.to_excel(writer, 'Result')
    writer.save()


if __name__ == "__main__":
    starting_time, deadline = input_judge_condition()
    calculation(starting_time, deadline)
