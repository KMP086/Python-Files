import urllib.parse
import pandas as pd
import pyodbc
import sqlalchemy
from sqlalchemy import *
import logging

def sqlread(drive, servername, dbname, uname, pword, query):
    logging.disable(logging.WARNING)
    sqlparam = urllib.parse.quote_plus(f"Driver={drive};Server={servername};Database={dbname};UID={uname};PWD={pword};")
    engine = sqlalchemy.create_engine("mssql+pyodbc:///?odbc_connect={}".format(sqlparam), echo=True)
    conn = engine.connect()
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    resultset = pd.read_sql(sqlalchemy.text(query), conn)
    return(resultset)
