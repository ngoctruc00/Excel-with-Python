# khai báo thư viện
import pandas as pd
import xlwings as xw


def get_value(row,column,excel_path):
    df= pd.read_excel(excel_path, 0, header=None)
    value = str(df.iloc[row, column])
    return value

def check(diemdanh_path, namdau_path, count):
    # read data
    df_diemdanh = pd.read_excel(diemdanh_path, 0, header=None)
    df_namdau = pd.read_excel(namdau_path, 0, header=None)
    # write data
    wb = xw.Book(diemdanh_path).sheets[1]

    """process data"""
    for j in range(df_namdau.shape[0]):
        mssv = get_value(j, 0,namdau_path)
        print(mssv)
        for i in range(int(df_diemdanh.shape[0])):
            if i != 0:
                value = str(df_diemdanh.iloc[i, 4])
                if value == mssv:
                    count = count + 1
                    name = "A" + str(count)
                    mssv = "B" + str(count)
                    wb.range(str(name)).value = df_diemdanh.iloc[i, 2]
                    wb.range(str(mssv)).value = df_diemdanh.iloc[i, 4]

    print('soluong la:',count)

    return count
if __name__=="__main__":

    diemdanh_path ='DIEMDANH.xls'
    namdau_path = 'MSSVcheck.xls'
    count = 0
    print(check(diemdanh_path,namdau_path,count))
    mb = xw.Book(diemdanh_path)
    mb.save()
    print('done save')
