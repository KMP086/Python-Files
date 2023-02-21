#pip install pyodbc
import pyodbc
#s for single data or m for multiple data or n for no display
def readsql(drive,servername,dbname, uname, pword, query, smn):
    conscript = "Driver=" + drive + ";Server=" + servername + ";Database=" + dbname + ";UID=" + uname + ";PWD=" + pword
    print(conscript)
    cndb = pyodbc.connect(conscript)
    cursor = cndb.cursor()
    #stored proc query = 'exec sp_sproc(123, 'abc')'
    cursor.execute(query)
    if smn != 'n':
        for item in cursor:
            if smn == 's':
                return item[0]
            elif smn == 'm':
                return item
    elif smn == 'n':
        return "data has been altered"

