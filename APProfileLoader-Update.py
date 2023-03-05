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
    fpath = 'C:\\Users\\Kim.Pambid\\OneDrive - DSV\\Desktop\CW1 Mena\\AP Profile CW1\\'
    fcount = folderfiles(fpath, 'c', 0)
    ir = []
    ic = []
    ih = []
    p = 0
    val = []
    hrf = 0 #row that will be set as column name
    sref = 2 #Starting df row/ reference row
    sqlcred = ('{SQL Server}', 'Server Name', 'Data Base', 'USER NAME', 'PASSWORD') # SQl Credentials
    #No of headers to display////////////////////////////////////////////////////////////////////////
    cquery = 'select count(HeaderName) from [SQL_MNLDB].[dbo].[tblAPHeaderRef]'
    csql = readsql(*sqlcred, cquery, 's')
    #////////////////////////////////////////////////////////////////////////////////////////////////

    for f in range(fcount):
        fname = folderfiles(fpath, 'd', f)

        for cl in range(int(csql)):
            query = 'Select HeaderName from SQL_MNLDB.dbo.tblAPHeaderRef where ID =' + str(cl + 1)
            sitem = readsql(*sqlcred, query, 's').strip()
            #print(sitem, len(sitem))
            sht = excelhloc(fpath, fname, 0, sitem, 's', 0)
            l = excelhloc(fpath, fname, 0, sitem, 'l', 0)
            r = excelhloc(fpath, fname, 0, sitem, 'h', 0)
            if r != None:
                c = excelhloc(fpath, fname, 0, sitem, 'c', r)
                hn = excelitem(fpath, fname, sht, r, c)
            #array input
            if r != None:
                ir.insert(cl, r-2)
                ic.insert(cl, c-2)
                ih.insert(cl, hn)
            elif r == None:
                ih.insert(cl, 'Unnamed: ' + str(cl+1))
        # bulk sql parameter
        # insert into query
        fn = str(fname)
        tdate = str(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        #Pandas DataFrame////////////////////////////////////////////////////////////////////////////
        pd.options.display.max_columns = None
        pd.options.display.max_rows = None
        #print(ir[0])
        df = pd.read_excel((fpath + fname), sheet_name=sht, usecols="B:DD", skiprows=ir[0], header=0)
        df.columns = df.iloc[0]
        #print(df)
        #header Change///////////////////////////////////////////////////////////////////////////////
        w = 0
        for h in range(int(csql)):
                hq = 'Select QueryHeader from SQL_MNLDB.dbo.tblAPHeaderRef where ID =' + str(h + 1)
                hitem = readsql(*sqlcred, hq, 's').strip()
                #print(ih[h])
                if ih[h] == None:
                    w = w + 1
                    dfheader = 'Unnamed: ' + str(w)
                    df.columns = df.columns.str.replace(dfheader, hitem)
                elif ih[h] != None:
                    df.columns = df.columns.str.replace(ih[h], hitem)
        print(df.iloc[sref:3, ic[0]:csql])
        #Fields need to be inserted to SQl/////////////////////////////////////////////////////////////
        #Indicate here what columns that u need//////////////////////////////////////////////////////////////////////////
        setdf = df.iloc[sref:l, ic[0]:csql]
        setdf.insert(0, 'DB_Date', tdate)
        setdf.insert(1, 'F_Name', fn)
        #print(records)
        sqlcol = ['DB_Date', 'F_Name', 'OrgCode', 'OrgName', 'Port', 'Grouping', 'Relation', 'Consol',
                  'Curr', 'BankAcc', 'ChargeCode', 'SettleGrp', 'CrLimit', 'PmtTerm', 'PmtDays',
                  'WHTTax', 'PayInv', 'QualAssure']
        #////////////////////////////////////////////////////////////////////////////////////////////
        p = math.trunc(l/500)
        if p == 0:
            records = setdf[sqlcol].iloc[sref:l,]
            bulksql(*sqlcred, records, 'tblAPProfile', 'append')
        if p >= 0:
            for y in range(p):
                v = 502 * (y + 1)
                if y == 0:
                    s = 2
                    v = 502
                records = setdf[sqlcol].iloc[s:v,]
                bulksql(*sqlcred,  records, 'tblAPProfile', 'append')
                s = v
                i = l - v

                if i < 502:
                    v = l
                    records = setdf[sqlcol].iloc[s:v,]
                    bulksql(*sqlcred, records, 'tblAPProfile', 'append')

    print("Process Complete!!!")
