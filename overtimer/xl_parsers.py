import pandas as pd

def personal_parser(filename):
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