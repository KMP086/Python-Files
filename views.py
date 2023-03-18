from flask import Blueprint, redirect, url_for, render_template, request, jsonify
from SQLConnection import *
from ExcelToSQLApp import approfile
from pandas import *
import pandas as pd
#https://www.youtube.com/watch?v=9MHYHgh4jYc
#Activate setting of parameters and argruments, plus route
#Website Front end and HTML CSS Display
query = str("Select Distinct ReportName from Database.dbo.Table")
dquery = str("SELECT *  FROM Database.dbo.Table")
sqlcred = ('{SQL Server}', 'ServerName', 'Database', 'UserName', 'Password') # SQl Credentials
display = Blueprint(__name__, "ExcelToSQL")

# render/ display url
@display.route("/")
def home():
    #dropdown list///////////////////////////////
    qlist = readsql(*sqlcred, query, 'm')
    #table///////////////////////////////////////
    ilist = bulkdisql(*sqlcred, dquery)
    loc = []
    loc.insert(0, len(ilist.columns)) #columns
    loc.insert(1, len(ilist)) #row
    itable = ilist.iloc[0:(loc[1]), 0:loc[0]]
    #itable.to_html convert to html
    return render_template("index.html", ql=qlist, tables=[itable.to_html(index=False, header=True, table_id="dtbl")])

# get the value from url////////////////////////////////////////
@display.route("/", methods=['POST','GET'])
def hometbox():
    if request.method == "POST":
        # get item from input/textbox
        i = request.form['ii']
        # get item from select/dropdownlist
        p = request.form['ptype']
        print(i, str(p))
        try:
            if i != "":
                approfile(i, p)
                return home()
            elif i != "": return home()
        except: return home()
    else:
        return home()



