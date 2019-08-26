#!/usr/bin/env python3
import pandas as pd
import datetime


"""
help to calculate to do number between starting time and deadline of different people
"""


def input_judge_condition():
    starting_time = input('input starting time(eg.2019-07-31):')
    starting_time = datetime.datetime.strptime(starting_time, '%Y-%m-%d').date()
    deadline = input('input deadline (eg.2019-09-01):')
    deadline = datetime.datetime.strptime(deadline, '%Y-%m-%d').date()
    return starting_time, deadline


def calculate_to_do_number(starting_time, deadline):
    df = pd.read_excel("input.xlsm", sheet_name='Sheet2', header=0,
                       names=['SQ', 'PSQ', '0', '1', '2', '3', '4', '5'])
    df1 = pd.DataFrame()
    df['8月数量'] = 0
    for i in ['0', '1', '2', '3', '4', '5']:
        if ":" in df[i]:
            try:
                df1[i] = pd.to_datetime(df[i], format="%Y-%m-%d %H:%M:%S")
            except ValueError:
                print('Column' + i + '\'s format is not supported! ')
            else:
                df['8月数量'] += df1[i].apply(lambda x: 1 if deadline > x.date() > starting_time else 0)
        else:
            try:
                df1[i] = pd.to_datetime(df[i], format="%Y-%m-%d")
            except ValueError:
                print('Column' + i + '\'s format is not supported! ')
            else:
                df['8月数量'] += df1[i].apply(lambda x: 1 if deadline > x.date() > starting_time else 0)
    writer = pd.ExcelWriter('output.xlsx')
    df.to_excel(writer, 'Sheet1', index=False)
    grouped = df['8月数量'].groupby(df['SQ']).sum()
    grouped.to_excel(writer, 'Result')
    writer.save()


if __name__ == "__main__":
    starting_time, deadline = input_judge_condition()
    calculate_to_do_number(starting_time, deadline)
