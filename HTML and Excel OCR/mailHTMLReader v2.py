import pandas as pd
import numpy as np
import re
from sqlodbc import *
import json
import sys
from bs4 import BeautifulSoup
import lxml
import html5lib
import warnings
warnings.filterwarnings("ignore")


jsonitems = [] #private declaration

def listrecords(itemheader, itemdf, records, acct):
      sqlcred = ('*******', '*******', '*******', '*******', '*******')
      if itemheader == 'clientname':
            try:
                  if email == "": clientid = "Null"
                  else:
                      clientquery = "SELECT id FROM clients WHERE *****.dbo.clients.name = '" + str(email) + "'"
                      clientid = sqlread(*sqlcred, clientquery)
            except: clientid = "Null"
            jsonitems.insert(int(records), {'client_id': clientid})
      elif itemheader != 'clientname':
            itempara = []
            itempara = str(itemdf).split('x')
            if acct == 'Nestle' and itemheader == 'length': jsonitems.insert(int(records), {itemheader: str(itempara[0])})
            elif acct == 'Nestle' and itemheader == 'width': jsonitems.insert(int(records), {itemheader: str(itempara[1])})
            elif acct == 'Nestle' and itemheader == 'height': jsonitems.insert(int(records), {itemheader: str(itempara[2])})
            else: jsonitems.insert(int(records), {itemheader: itemdf})
      #print(jsonitems)

#left to right reading!!!!
def search(df: pd.DataFrame, substring: str, case: bool = False) -> pd.DataFrame:
      mask = np.column_stack([df[col].astype(str).str.contains(substring.lower(), case=case, na=False) for col in df])
      return df.loc[mask.any(axis=1)]

def ReadHtml():
      sqlcred = ('*******', '*******', '*******', '*******', '*******')
      mailid = str(sys.argv[1])
      #mailid = 27
      sqldetailquery = "exec [*****].[dbo].[getmaildetails] @mailheader = '" + str(mailid) + "'"
      sqldetailheaders = sqlread(*sqlcred, sqldetailquery)

      #print(sqldetailheaders)
      for sqlcntr in range(len(sqldetailheaders['alias'])):
            sqlloc = sqldetailheaders['alias'].loc[int(sqlcntr)]
            acct = sqldetailheaders['account'].loc[int(sqlcntr)]
            email = sqldetailheaders['email'].loc[int(sqlcntr)]

            #print(sqlloc)
            fileloc = 'C:\\*********\\wwwroot\\************\\storage\\app\\public\\tools\\emailfiles\\' + sqlloc
            
            #print(fileloc)
            
            #Search Item//////////////////////////////////////////////////////
            sqlquery = str("select ref_value, additional_row, additional_column, ref_key from ******.dbo.mail_html_coordinates where account_ref = '" + str(acct) + "'")
            sqlheaders = sqlread(*sqlcred, sqlquery)
            #print(sqlheaders)
            
            #fpath = str(fileloc.replace('\\', '/')).strip() + '/' + str(mailid) + '/'
            fpath = fileloc + '\\' + str(mailid) + '\\'
            #fpath = 'D:\\Project\\envpython\\env\\testfiles\\'
            
            #print(fpath)
            #print(fpath + 'index.html')
            pd.options.display.max_columns = None  # max display from print columns
            pd.options.display.max_rows = None  # max display from print rows
            #clientname = str(sqlheaders['client'].loc[0])
            # read excel & set as new df/////////////////////////////////
            df = pd.concat(pd.read_html(fpath + 'index.html'), axis=1)
            df.columns = df.iloc[0]
            #print(df)
            #format text & remove special characters////////////////////////////////////////////////////
            charac = ['()?*']
            try:
                try:
                     df.replace('\n', '', regex=True, inplace=True)  # remove new line
                     dfjson = df.to_json(orient="columns")
                except: dfjson = df.to_json(orient="columns")
                try: formatjson = re.sub(str(charac), '', dfjson)
                except: formatjson = dfjson
                newdf = pd.read_json(formatjson)
            except: newdf = df
            #print(newdf)
            # ///////////////////////////////////////////////////////////////////////////////////////////

            for records in range(len(sqlheaders)):
                item = sqlheaders.iloc[int(records), 0].strip()
                rown = int(sqlheaders.iloc[int(records), 1]) #SQL Header row
                coln = int(sqlheaders.iloc[int(records), 2]) #SQL Header col
                itemheader = sqlheaders.iloc[int(records), 3].strip()
                if item != 'Null':
                   try:
                       SearchDF = pd.DataFrame(search(newdf, item))
                       if rown == 0 and coln >= 1:
                          counter = 0
                          #print(len(SearchDF.columns))
                          for counter in range(len(SearchDF.columns)):
                              itemdetail = str(SearchDF.iloc[0, int(counter)])
                              #print(itemdetail)
                              if itemdetail == item:
                                 colnum = int(counter)
                                 break
                          itemdf = SearchDF.iloc[0 + rown, colnum + coln]
                          listrecords(itemheader, itemdf, records, acct)

                       elif rown >= 1 and coln == 0:
                            counter = 0
                            #might have an error/////////////////////////////////////////////
                            rowvalue = str(SearchDF.index.values).replace('[', '').strip()
                            #////////////////////////////////////////////////////////////////
                            itemrow = int(rowvalue[0:rowvalue.find(' ')].strip())
                            for counter in range(len(SearchDF.columns)):
                                itemdetail = str(SearchDF.iloc[0, int(counter)])
                                #print(itemdetail)
                                if itemdetail == item:
                                   colnum = int(counter)
                                   break
                            itemdf = df.iloc[itemrow + rown, colnum + coln]
                            listrecords(itemheader, itemdf, records, acct)
                   except:
                       if itemheader == 'clientname': jsonitems.insert(int(records), {'client_id': "NULL"})
                       elif itemheader != 'clientname': jsonitems.insert(int(records), {itemheader: "NULL"})

                elif item == "Null":
                     if itemheader == 'clientname': jsonitems.insert(int(records), {'client_id': "NULL"})
                     elif itemheader != 'clientname': jsonitems.insert(int(records), {itemheader: "NULL"})
                     continue

            #print(jsonitems)
            #/////////////////////////////////////////////////////////////////
            # df.to_json convert to json with
            result = json.dumps(jsonitems)
            jsondata = re.sub('[{}]', '', result)
            return jsondata




