import win32com.client
import os.path
import json
Def getDirectReportee(name):
  outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
  filename = 'data_dr.json'
  dd = {}
  if os.path.isfile(filename):
    with open(filename,'r') as for:
      dd = json.load(fp)
  if name in dd:
    return dd[name]
  else :

