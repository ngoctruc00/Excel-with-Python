import pandas as pd
import xlwings as xw

excel_file = 'DIEMDANH.xls'
excel_path= 'MSSVcheck.xls'

wb1 = xw.Book('DIEMDANH.xls').sheets[1]

df = pd.read_excel(excel_file, 0, header=None)
df_mssv = pd.read_excel(excel_path, 0, header=None)

def get_value(row,column):
    value = str(df_mssv.iloc[row, column])
    return value

count =  0
for j in range(df_mssv.shape[0]):
    mssv = get_value(j,0)
    print(mssv)
    for i in range(int(df.shape[0])):
        if i != 0:
            value = str(df.iloc[i, 4])
            if value == mssv:
                count = count + 1
                name = "A" + str(count)
                mssv = "B" + str(count)
                wb1.range(str(name)).value = df.iloc[i, 2]
                wb1.range(str(mssv)).value = df.iloc[i, 4]
                print('so luong la ',count)

