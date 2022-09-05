import pandas as pd
import openpyxl

path = './vocab_All v2.xlsx'

def get_frequency_and_margin(path):
    xls = pd.ExcelFile(path)
    results = ""
    max_freqs = []
    min_freqs = []

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet)
        # get frequency
        x = df['frequency'].values
        max_freq = x.max()
        min_freq = x.min()
        # get margin
        common = f'{max_freq*0.6:.0f} to {max_freq}'
        normal = f'{max_freq*0.25+1:.0f} to {max_freq*0.6-1:.0f}'
        rare = f'{min_freq} to {max_freq*0.25:.0f}'
        # get result
        result = f'{sheet}: Min:{min_freq}, Max:{max_freq}, Common:{common}, Normal:{normal}, Rare:{rare}\n'
        results = results + result
        max_freqs.append(max_freq)
        min_freqs.append(min_freq)

    return results, max_freqs, min_freqs

def record_frequency_and_margin(path):
    results, max_freq, min_freq = get_frequency_and_margin(path)
    with open('distribution and margin.txt', 'w') as f:
        f.write(results)

def classify(path):
    results, max_freqs, min_freqs = get_frequency_and_margin(path)
    wb = openpyxl.load_workbook(path)
    count = 0
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        ws.auto_filter.ref(filterColumn=("C"))
        # # classify classes
        # for i in range(ws.max_row-1):
        #     freq = ws["C"+str(i+2)].value
        #     if freq >= 0.6*max_freqs[count]:
        #         # common
        #         ws["B"+str(i+2)].value = 1
        #     elif freq <= 0.25*max_freqs[count]:
        #         # rare
        #         ws["B"+str(i+2)].value = 3
        #     else:
        #         # normal
        #         ws["B"+str(i+2)].value = 2
        # count = count + 1

    wb.save(path)

xls = pd.ExcelFile(path)
for sheet in xls.sheet_names:
    df = pd.read_excel(xls, sheet)
    # in small to large order
    df = df.sort_values(by='frequency')
    for i in range(len(df)):
        if i <= len(df)*0.25:
            df.at[i, 'level'] = 1
        elif i >= len(df)*0.6:
            df.at[i, 'level'] = 3
        else:
            df.at[i, 'level'] = 2
    df.to_excel(xls, sheet_name=sheet)
