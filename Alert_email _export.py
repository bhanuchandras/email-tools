import win32com.client
from datetime import datetime
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)

def Processmails(f):
  col_sep="~"
  for i in f.Items:
    if ( i.ReceivedTime.month == 11 and i.ReceivedTime.month == 25 ) :
      sev = i.Subject[:1]
      global c
      filehandler.writer(
         str(.....))
      c+=1

startTime = datetime.now()
index_dict={}
c=1

