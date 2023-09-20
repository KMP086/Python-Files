import win32com.client as win32
import os
from sqlodbc import *
import sys

class outlookmail():
      try:
            mailid = str(sys.argv[1])
            #mailid = 27
            #Outlook recipient & text////////////////////////////////////
            sqlcred = ('*******', '*******', '*******', '*******', '*******')
            mailquery = "select sender, subject, client from SQT_QA.dbo.mail_headers where id = " + str(mailid) + ""
            sqlmail = sqlread(*sqlcred, mailquery)
            #parts////
            recipient = sqlmail['sender'].loc[0]
            #recipient = sqlmail['client'].loc[0]
            subject = sqlmail['subject'].loc[0]
            text = 'This is a Test!!!'

            outlook = win32.Dispatch('outlook.application')
            mail = outlook.CreateItem(0)
            mail.To = recipient
            mail.Subject = subject
            mail.HtmlBody = text
            #attachment/////////////////////////////
            #attachloc = os.getcwd() +"\\file.ini"
            #mail.Attachments.Add(attachloc)
            mail.Display(True)
      except:
            print('Please close Outlook draft!!!')