#reference https://github.com/O365/python-o365#authentication
#pip install o365
#Search for: 'Source Root' then Add it!!!
#/////////////////////////////////////////////////////////////////////////////////////////////
#Access Authentication:
from O365 import Account
def auth(clientid, clientsecret, tenantid):
    credentials = (str(clientid), str(clientsecret))
    acct = Account(credentials, auth_flow_type='credentials', tenant_id=str(tenantid))
    if acct.authenticate():
        print('Authenticated!')

#/////////////////////////////////////////////////////////////////////////////////////////////
#Get Workbook: --> Authenticate it first
from O365.excel import WorkBook
fn = 'Month End Exchange Rates 0123.xlsx'