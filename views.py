from flask import Blueprint, redirect, url_for, render_template, request, jsonify
import datetime
from SQLConnection import *
from ExcelToSQLApp import approfile
from FOREXLoadApp import *
from pandas import *
import pandas as pd
#https://www.youtube.com/watch?v=9MHYHgh4jYc
#Activate setting of parameters and argruments, plus route
#Website Front end and HTML CSS Display
query = str("Select Distinct ReportName from Database.dbo.Table")
sqlcred = ('Driver', 'Server', 'Database', 'UID', 'PWD') # SQl Credentials
display = Blueprint(__name__, "ExcelToSQL")
# render/ display url
@display.route("/ExceltoSQL")
def home(ptype=None):
    #dropdown list///////////////////////////////
    qlist = readsql(*sqlcred, query, 'm')

    if ptype == None: ttbl = "Database.dbo.Table"
    else:
        tquery = str("Select Distinct DBTableName from Database.dbo.Table Where ReportName ='" + str(ptype) + "'")
        ttbl = readsql(*sqlcred, tquery, 's')
    #Date today//////////////////////////////////
    dt = datetime.datetime.now()
    dt = dt.year - 2015
    #table///////////////////////////////////////
    dquery = str("SELECT *  FROM " + str(ttbl))
    print(dquery)
    ilist = bulkdisql(*sqlcred, dquery)
    loc = []
    loc.insert(0, len(ilist.columns)) #columns
    loc.insert(1, len(ilist)) #row
    itable = ilist.iloc[0:(loc[1]), 0:loc[0]]
    #itable.to_html convert to html with(<table><tr><td> included)
    return render_template("index.html", ql=qlist, tables=[itable.to_html(index=False, header=True, table_id="dtbl")], rtype=ptype, tdate=dt)
    #itable.to_json convert to json with
    #return render_template("index.html", ql=qlist, tables=itable.to_json(index=False, orient='split'))


# get the value from url////////////////////////////////////////
@display.route("/ExceltoSQL", methods=['POST','GET'])
def hometbox():
    if request.method == "POST":
        # get item from input/textbox
        i = request.form['ii']
        # get item from select/dropdownlist to insert
        p = request.form['ptype']
        try:
            if i != "":
                if p == 'Forex':
                    appForex(i)
                    return home(p)
                else:
                    approfile(i, p)
                    return home(p)
            elif i == "":
                if p == 'Forex': return home(p)
                else: return home(p)
        except: return home(p)
    else:
        return home(request.form['ptype'])



