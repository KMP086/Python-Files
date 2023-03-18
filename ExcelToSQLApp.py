#https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.iloc.html
# pip install datetime
# pip install pandas
import math
from SQLConnection import *
from AccessFiles import *
import datetime
import pandas as pd
import re

def approfile(floc, ptype):
    #Setup Variables//////////////////////////////////////////////////////////////////////////////////////////
    fpath = str(floc.replace('\\', '/')).strip() + '\\'
    hrf = 0 #row that will be set as column name
    sref = 2 #Starting df row/ reference row
    sqlcred = ('{SQL Server}', 'ServerName', 'Database', 'UserName', 'Password') # SQl Credentials
    sqlcolname = 'HeaderName'
    #Query will be used in the table//////////////////////////////////////////////////////////////////////////
    RepName = str(ptype)
    itemqry = str("Select QueryHeader from Database.dbo.Table Where ReportName = '" + RepName + "'")
    sqlcol = readsql(*sqlcred, itemqry, 'm')
    fsqlcol = []
    for scl in range(len(sqlcol)):
        fsqlcol.insert(scl, sqlcol[scl])
    fsqlcol.insert(0, 'DB_Date')
    fsqlcol.insert(1, 'F_Name')
    print(fsqlcol)
    #sqlcol = ['DB_Date', 'F_Name', 'OrgCode', 'OrgName', 'Port', 'Grouping', 'Relation', 'Consol',
    #          'Curr', 'BankAcc', 'ChargeCode', 'SettleGrp', 'CrLimit', 'PmtTerm', 'PmtDays',
    #          'WHTTax', 'PayInv', 'QualAssure']
    sqldatatbl = 'tblAPProfile'
    #Arrays////////////////////////////////////////////////////////////////////////////////////////////////////
    ir = []
    ic = []
    ih = []
    hitem = []
    qitem = []
    val = []
    specicol = [] #specified columns in data frame list
    #Get data  from folders/////////////////////////////////////////////////////////////////////////////////////
    p = 0 # initialize start for sql bulk loop
    fcount = folderfiles(fpath, 'c', 0)
    #//////////////////////////////////////////////////////////////////////////////////////////////////////////
    for f in range(fcount):
        fname = folderfiles(fpath, 'd', f)
        query = 'Select HeaderName from Database.dbo.Table where ID = 1'
        fitem = readsql(*sqlcred, query, 's').strip()
        sht = excelhloc(fpath, fname, 0, fitem, 's', 0)
        l = excelhloc(fpath, fname, 0, fitem, 'l', 0)
        r = excelhloc(fpath, fname, 0, fitem, 'h', 0)
        ir.insert(0, r - 2)
        #Pandas DataFrame -> Data Arrangement/////////////////////////////////////////////////////////////////////
        pd.options.display.max_columns = None
        pd.options.display.max_rows = None
        df = pd.read_excel((fpath + fname), sheet_name=sht, usecols="B:DD", skiprows=ir[0], header=0)
        df.columns = df.iloc[0]
        dfcnt = len(df.columns)
        hn = None

        for ch in range(int(dfcnt)):
            try:
                qquery = str(f"Select QueryHeader from Database.dbo.Table Where HeaderName = '{str(df.columns[int(ch)].strip())}'")
                qitem.insert(ch, readsql(*sqlcred, qquery, 's').strip())
                df.columns = df.columns.str.replace(df.columns[int(ch)], qitem[ch])
            except:
                qitem.insert(ch, 'None' + str(ch))
                df.columns = df.columns.str.replace(df.columns[int(ch)], qitem[ch])
                continue
        # bulk sql parameter/////////////////////////////////////////////////////////////////////////////////////
        fn = str(fname)
        tdate = str(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        #Fields need to be inserted to SQl///////////////////////////////////////////////////////////////////////
        setdf = df.iloc[sref:l, 0:dfcnt]
        setdf.insert(0, 'DB_Date', tdate)
        setdf.insert(1, 'F_Name', fn)
        #assert isinstance(setdf.reidex, object)
        # insert into query/////////////////////////////////////////////////////////////////////////////////////
        p = math.trunc(l/500)
        if p == 0:
            records = setdf[fsqlcol].iloc[sref:l,]
            bulksql(*sqlcred, records, sqldatatbl, 'append')
        if p >= 0:
            for y in range(p):
                v = 502 * (y + 1)
                if y == 0:
                    s = 2
                    v = 502
                records = setdf[fsqlcol].iloc[s:v,]
                print(records)
                bulksql(*sqlcred,  records, sqldatatbl, 'append')
                s = v
                i = l - v

                if i < 502:
                    v = l
                    records = setdf[fsqlcol].iloc[s:v,]
                    bulksql(*sqlcred, records, sqldatatbl, 'append')

    print("Process Complete!!!")
