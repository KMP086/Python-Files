# pip install datetime
from SQLConnection import *
from AccessFiles import *
import datetime


# connsql(drive,servername,dbname, uname, pword, query, smn)
class approfile:
    fpath = 'C:\\Users\\Kim.Pambid\\OneDrive - DSV\\Desktop\*************\\**************\\'
    fcount = folderfiles(fpath, 'c', 0)
    ir = []
    ic = []
    for f in range(fcount):
        fname = folderfiles(fpath, 'd', f + 1)

        for cl in range(2):
            query = 'Select HeaderName from SQL_MNLDB.dbo.tblAPHeaderRef where ID =' + str(cl + 1)
            sitem = readsql('{SQL Server}', '**********', '*********', '*********', '********', query, 's')
            r = excelhloc(fpath, fname, 0, sitem, 'h', 0)
            c = excelhloc(fpath, fname, 0, sitem, 'c', r)
            l = excelhloc(fpath, fname, 0, sitem, 'l', 0)
            sht = excelhloc(fpath, fname, 0, sitem, 's', 0)
            #array input
            ir.insert(cl, r)
            ic.insert(cl, c)

        # loop per
        for p in range(l):
            if p >= 1:
                ia = str(excelitem(fpath, fname, sht, ir[0] + p, ic[0]))
                ib = str(excelitem(fpath, fname, sht, ir[1] + p, ic[1]))
                if ia != 'None':
                    # insert into query
                    fn = str(fname)
                    tdate = str(datetime.datetime.now())
                    val = (tdate, fname, ia, ib)
                    iquery = "insert into SQL_MNLDB.dbo.tblAPProfile(DB_Date, F_Name, OrgCode, OrgName) values (?,?,?,?)"
                    altersql('{SQL Server}', '********', '********', '********', '*******', iquery, val)
                    print(val)