import pandas as pd
import numpy as np
import openpyxl
from sklearn import preprocessing

path = './vocab_All.xlsx'
xls = pd.ExcelFile(path)
wb = openpyxl.load_workbook(path)
for sheet in wb.sheetnames:
    ws = wb[sheet]

    df = pd.read_excel(xls, sheet)
    # get frequency
    x = df['frequency'].values.reshape(-1,1)

    # normalize frequency
    # min_max_scaler = preprocessing.MinMaxScaler()
    # x_scaled = min_max_scaler.fit_transform(x)
    # df2 = pd.DataFrame(x_scaled).values
    # ws["I1"].value = 'norm_freq'
    # for i in range(len(df2)):
    #     ws["I"+str(i+2)].value = df2[i][0]

    # classify classes
    ws["J1"].value = 'class'
    for i in range(ws.max_row-1):
        freq = ws["I"+str(i+2)].value
        if freq >= 0.6:
            ws["J"+str(i+2)].value = "common"
        elif freq <= 0.25:
            ws["J"+str(i+2)].value = "master"
        else:
            ws["J"+str(i+2)].value = "in-depth"


wb.save(path)

#     print(f'{sheet}: Min:{col.min()}, Max:{col.max()}')


