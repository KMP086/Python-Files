#https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.iloc.html
# pip install datetime
# pip install pandas
import math
from SQLConnection import *
from AccessFiles import *
import datetime
import pandas as pd


# connsql(drive,servername,dbname, uname, pword, query, smn)
class approfile:
    fpath = 'C:\\Users\\Kim.Pambid\\OneDrive\\********\\********\\'
    fcount = folderfiles(fpath, 'c', 0)
    ir = []
    ic = []
    p = 0
    val = []
    for f in range(fcount):
        fname = folderfiles(fpath, 'd', f)

        for cl in range(2):
            query = 'Select HeaderName from SQL_*****.dbo.tbl*********** where ID =' + str(cl + 1)
            sitem = readsql('{SQL Server}', '******', '*******', '*******', '*******', query, 's')
            r = excelhloc(fpath, fname, 0, sitem, 'h', 0)
            c = excelhloc(fpath, fname, 0, sitem, 'c', r)
            l = excelhloc(fpath, fname, 0, sitem, 'l', 0)
            sht = excelhloc(fpath, fname, 0, sitem, 's', 0)
            #array input
            ir.insert(cl, r-2)
            ic.insert(cl, c-2)

        # bulk sql parameter
        # insert into query
        fn = str(fname)
        tdate = str(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        #Pandas DataFrame//////////////////////////////////////////////////////////////////////////////////////
        df = pd.read_excel((fpath + fname), sheet_name=sht, usecols="B:DD", skiprows=ir[0], header=0)
        df.columns = df.columns.str.replace('Unnamed: 1', 'OrgCode')
        df.columns = df.columns.str.replace('Unnamed: 2', 'OrgName')

        p = math.trunc(l/500)
        if p == 0:
            records = df.iloc[2:l, [ic[0], ic[1]]]
            records.insert(0, 'DB_Date', tdate)
            records.insert(1, 'F_Name', fn)
            bulksql('{SQL Server}', '******', '******', '*******', '*******', records, '*******', 'append')
        if p >= 0:
            for y in range(p):
                v = 502 * (y + 1)
                if y == 0:
                    s = 2
                    v = 502
                records = df.iloc[s:v, [ic[0], ic[1]]]
                records.insert(0, 'DB_Date', tdate)
                records.insert(1, 'F_Name', fn)
                bulksql('{SQL Server}', '*******', '*******', '*******', '********',  records, '********', 'append')
                s = v
                i = l - v

                if i < 502:
                    v = l
                    records = df.iloc[s:v, [ic[0], ic[1]]]
                    records.insert(0, 'DB_Date', tdate)
                    records.insert(1, 'F_Name', fn)
                    bulksql('{SQL Server}', '******', '*******', '***********', '*******', records, '********', 'append')

    print("Process Complete!!!")
