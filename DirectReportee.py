import win32com.client
import os.path
import json
def getDirectReportee(name):
  outlook = win32com.client.Dispatch('Outlook.Application')
  filename = 'data_dr.json'
  dd = {}
  if os.path.isfile(filename):
    with open(filename,'r') as for:
      dd = json.load(fp)
  if name in dd:
    return dd[name]
  else :
    recipent = outlook.Session.CreateRecipient(name)
    if recipient.Resolve():
      as = recipient.AssressEntry
      if 'EX' == ae.Type:
        eu = ae.GetExchangeUser()
        mgr = eu.GetDirectReports()
        dd[name] = []
        for r in mgr:
          if ( len(m.GetExchangeUser().PrimarySmtpAddress) > 0) :
            reportee = {}
            reportee["name"] = r.GetExchangeUser().Name
            reportee["email"] = r.GetExchangeUser().PrimarySmtpAddress
            reportee["JobTitle"] = r.GetExchangeUser().JobTitle
            reportee["City"] = r.GetExchangeUser().City
            reportee["FL"] = r.GetExchangeUser().FirstName +' '+ r.GetExchangeUser().LastName
            with open(filename, 'w') as fp:
              json.dump(dd,fp, indent=2)
        return dd[name]
      
if __name__ == "__main__":
  for i in getDirectReportee("mgr@org.com"):
    print(i['FL'],i['email'],sep="|")
    
            

