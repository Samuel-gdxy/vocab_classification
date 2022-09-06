import pandas as pd

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
