from mailHTMLReaderv3 import *
from mailExcelReaderv3 import *
import re
import json
import jsonify

newlist = []
checkhtml = []
applyhtml = []
checkexcel = []
applyexcel = []


def jsontolist(recordlist, type):

    item = recordlist.split(',')
    for c in range(len(item)):
        hitem = item[c].replace("[", "").replace("]", "")
        hitem = hitem.replace('"', '').replace('"', '')
        hitem = re.sub(' +', '', hitem)
        sitem = 0
        litem = 0
        sitem = int(hitem.find(':')) + 1
        rcount = int(hitem.find(':'))
        litem = int(len(hitem))
        scount = litem - sitem
        header = (hitem[:rcount].strip())
        word = (hitem[-scount:].strip())
        if type == 'html':
            checkhtml.insert(c, word)
            applyhtml.insert(c, {header: word})
        elif type == 'excel':
            checkexcel.insert(c, word)
            applyexcel.insert(c, {header: word})


class jsonOCR():

      try:
            jsontolist(ReadHtml(), 'html')
            jsontolist(ReadExcel(), 'excel')

            for w in range(17):
                if checkhtml[w] == 'NULL': newlist.insert(w, applyexcel[w])
                elif checkhtml[w] != 'NULL' and checkexcel[w] != 'NULL': newlist.insert(w, applyhtml[w])
                else: newlist.insert(w, applyhtml[w])

            #result = json.dumps(newlist)
            #jsondata = re.sub('[{}]', '', result)
            jsondata = json.dumps(newlist)


            #print(jsondata)
      except:
            try: jsondata = ReadHtml()
            except: jsondata = ReadExcel()
      def jsonresults(rvalues=jsondata):
          jresult = rvalues
          return jresult
print(jsonOCR.jsonresults())


