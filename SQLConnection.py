#upgrade -m pip install --upgrade 'sqlalchemy<2.0'
#pip install pyodbc
#pip install sqlalchemy
# pip install pandas

import urllib.parse
import pandas as pd
import pyodbc
import sqlalchemy
from sqlalchemy import *

#s for single data or m for multiple data or n for no display

def readsql(drive,servername,dbname, uname, pword, query, sm):
    w = 0
    conscript = "Driver=" + drive + ";Server=" + servername + ";Database=" + dbname + ";UID=" + uname + ";PWD=" + pword
    #print(conscript)
    cndb = pyodbc.connect(conscript)
    cursor = cndb.cursor()
    #stored proc query = 'exec sp_sproc(123, 'abc')'
    result = cursor.execute(query)
    if sm == 's':
        for item in result:
            #print(item)
            return item[0]
    elif sm == 'm':
        for row in result.fetchall():
            w = w + 1
            if w == 1:
                i = [str(row[0]).strip()]
            elif w > 1:
                i.insert(w, str(row[0]).strip())
        #print(i)
        return i

#for insert into or update set (single entry only)
def altersql(drive,servername,dbname, uname, pword, query, val):
    #query = 'insert into table(a,b,c) value(?,?,?)'
    #values = [(a,b,c)]
    conscript = "Driver=" + drive + ";Server=" + servername + ";Database=" + dbname + ";UID=" + uname + ";PWD=" + pword
    #print(conscript)
    cndb = pyodbc.connect(conscript)
    cursor = cndb.cursor()
    cursor.execute(query, val)
    cndb.commit()

#insert and update multiple rows
# sqltype type append or replace
#note dataframe column name must be identical SQL Column Name
def bulksql(drive,servername,dbname, uname, pword, dataf, sqltable, sqltype):
    sqlparam = urllib.parse.quote_plus(f"Driver={drive};Server={servername};Database={dbname};UID={uname};PWD={pword}")
    engine = create_engine("mssql+pyodbc:///?odbc_connect={}".format(sqlparam), use_setinputsizes=False)
    pd.set_option('display.max_columns', None)
    df = pd.DataFrame(dataf)
    df = df.astype(str)
    print(df)
    df.to_sql(con=engine, name=sqltable, if_exists=sqltype, index=False, chunksize=1000)


#bulk display select from where
def bulkdisql(drive, servername, dbname, uname, pword, sqlquery):
    sqlparam = urllib.parse.quote_plus(f"Driver={drive};Server={servername};Database={dbname};UID={uname};PWD={pword};")
    engine = sqlalchemy.create_engine("mssql+pyodbc:///?odbc_connect={}".format(sqlparam), echo=True)
    conn = engine.connect()
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    resultset = pd.read_sql(sqlalchemy.text(sqlquery), conn)
    #print(resultset)
    return(resultset)

