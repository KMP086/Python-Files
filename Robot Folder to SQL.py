# pip install datetime
from SQLConnection import *
from AccessFiles import *
import datetime


# connsql(drive,servername,dbname, uname, pword, query, smn)
class approfile:
    fpath = 'C:\\Users\\*******\\'
    fcount = folderfiles(fpath, 'c', 0)
    ir = []
    ic = []
    for f in range(fcount):
        fname = folderfiles(fpath, 'd', f + 1)

        for cl in range(2):
            query = 'Select HeaderName from SQL_*****.dbo.tblAPHeaderRef where ID =' + str(cl + 1)
            sitem = readsql('{SQL Server}', '*****', 'SQL_*****', '****Developer$', '******', query, 's')
            sht = 'Template'
            r = excelhloc(fpath, fname, sht, sitem, 'h', 0)
            c = excelhloc(fpath, fname, sht, sitem, 'c', r)
            l = excelhloc(fpath, fname, sht, sitem, 'l', 0)
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
                    altersql('{SQL Server}', '******', 'SQL_*****', '*****Developer$', '*****', iquery, val)
