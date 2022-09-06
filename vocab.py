import pandas as pd
import openpyxl


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

# sort excel rows by frequency
def sort(path):
    xls = pd.ExcelFile(path)
    dfs = []
    sheets = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet)
        # in small to large order
        df = df.sort_values(by='frequency')
        dfs.append(df)
        sheets.append(sheet)
    for i in range(len(sheets)):
        df = dfs[i]
        sheet = sheets[i]
        # save dataframe to excel file
        with pd.ExcelWriter(xls, mode='a', if_sheet_exists='replace', engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)

# classify vocab level
def classify(path):
    xls = pd.ExcelFile(path)
    dfs = []
    sheets = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet)
        rare_max = round(len(df) * 0.25)
        common_max = round(len(df) * 0.6)
        for index in df.index:
            if index <= rare_max:
                df.at[index, 'level'] = 1
            elif index >= common_max:
                df.at[index, 'level'] = 3
            else:
                df.at[index, 'level'] = 2
        df = df.sort_values(by='vocab id')
        dfs.append(df)
        sheets.append(sheet)
    for i in range(len(sheets)):
        df = dfs[i]
        sheet = sheets[i]
        # save dataframe to excel file
        with pd.ExcelWriter(xls, mode='a', if_sheet_exists='replace', engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=sheet, index=False)

path = './vocab_All v2.xlsx'
classify(path)
