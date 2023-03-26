#https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.iloc.html
# pip install datetime
# pip install pandas
import math
from SQLConnection import *
from AccessFiles import *
import datetime
import pandas as pd

# connsql(drive,servername,dbname, uname, pword, query, smn)
def appForex(floc):
    sqlcred = ('Driver', 'Server', 'Database', 'UID', 'PWD') # SQl Credentials
    fpath = str(floc.replace('\\', '/')).strip() + '\\'
    fcount = folderfiles(fpath, 'c', 0)
    p = 0
    val = []
    for f in range(fcount):
        fname = folderfiles(fpath, 'd', f)

        for cl in range(2):
            r = 8 #Forex Start Row
            c = 2 #Forex Start Column
            l = excelhloc(fpath, fname, 0, '', 'l', 0)
            sht = excelhloc(fpath, fname, 0, '', 's', 2)

        # bulk sql parameter
        # insert into query
        fn = str(fname)
        tdate = str(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        #Pandas DataFrame//////////////////////////////////////////////////////////////////////////////////////
        pd.set_option('display.max_columns', None)
        df = pd.read_excel((fpath + fname), sheet_name=sht, usecols="B:E", skiprows=6, header=0)
        df.columns = df.columns.str.replace('Unnamed: 1', 'Ctry_Name')
        df.columns = df.columns.str.replace('Unnamed: 2', 'Curr')
        df.columns = df.columns.str.replace('Unnamed: 3', 'CurrCode')
        df.columns = df.columns.str.replace('Unnamed: 4', 'F_Amt')

        p = math.trunc(l/500)
        if p == 0:
            records = df.iloc[0:l, ]
            records.insert(0, 'DB_Date', tdate)
            records.insert(1, 'F_Name', fn)
            bulksql(*sqlcred, records, 'tblAPFOREX', 'append')
        if p >= 0:
            for y in range(p):
                v = 502 * (y + 1)
                if y == 0:
                    s = 2
                    v = 502
                records = df.iloc[s:v, ]
                records.insert(0, 'DB_Date', tdate)
                records.insert(1, 'F_Name', fn)
                bulksql(*sqlcred,  records, 'tblAPFOREX', 'append')
                s = v
                i = l - v

                if i < 502:
                    v = l
                    records = df.iloc[s:v, ]
                    records.insert(0, 'DB_Date', tdate)
                    records.insert(1, 'F_Name', fn)
                    bulksql(*sqlcred, records, 'tblAPFOREX', 'append')

    print("Process Complete!!!")
