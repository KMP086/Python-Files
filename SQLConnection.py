#pip install pyodbc
import pyodbc
import pandas as pd
#s for single data or m for multiple data or n for no display
def readsql(drive,servername,dbname, uname, pword, query, sm):
    conscript = "Driver=" + drive + ";Server=" + servername + ";Database=" + dbname + ";UID=" + uname + ";PWD=" + pword
    #print(conscript)
    cndb = pyodbc.connect(conscript)
    cursor = cndb.cursor()
    #stored proc query = 'exec sp_sproc(123, 'abc')'
    cursor.execute(query)
    for item in cursor:
        if sm == 's':
            return item[0]
        elif sm == 'm':
            return item

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