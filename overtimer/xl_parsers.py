import pandas as pd
import datetime


def get_personal_df(filename):
    data = pd.ExcelFile(filename).parse(sheet_name=0, skiprows=13)
    data.columns = list('ABCDEFGHIJKLMNOP')
    data.drop(
        ['A', 'B', 'D', 'E', 'I', 'J', 'F', 'G', 'L', 'M', 'N', 'O', 'P'],
        axis='columns',
        inplace=True,
    )
    data.columns = ['names', 'numbers', 'hours']

    data.dropna(inplace=True)
    data = data[data['names'] != "3"]

    data['hours'] = data['hours'].apply(int)
    data['numbers'] = data['numbers'].apply(int)
    return data


def to_zero_date(item):
    if not isinstance(item, datetime.datetime):
        return datetime.date(1900, 1, 1)
    return item


def get_order_number(filename):
    sheets = pd.ExcelFile(filename)
    frames = []
    for sheet_name in sheets.sheet_names:
        sheet = sheets.parse(sheet_name)
        sheet.dropna(inplace=True)
        if sheet.shape[0] and sheet.shape[1] == 3:
            sheet.drop(sheet.columns[[0]], axis=1, inplace=True)
            sheet.columns = ['numbers', 'dates'] if '/' in str(sheet[sheet.columns[[0]]].values[0]) else ['dates', 'numbers']
            frames.append(sheet)
    data = pd.concat(frames)
    data['dates'] = data['dates'].apply(to_zero_date)
    return data.sort_values(by='dates').iloc[-1:]['numbers'].values[0]
