from calendar import weekday
from email import message
from fileinput import filename
from openpyxl import Workbook
import requests
requests.packages.urllib3.disable_warnings()
import os
import datetime
from requests_kerberos import HTTPKerberosAuth, REQUIRED, OPTIONAL
import openpyxl
from openpyxl.workbook.protection import WorkbookProtection
import time
import datetime as dt
from datetime import timedelta
from datetime import date
import datetime
import csv




def testbrowser():
    try:
        url = 'https://maxis-service-prod-pdx.amazon.com/issues?q=containingFolder:(82c69b17-7c44-48c1-8374-7f4850747021)'
        response = requests.get(url, auth=HTTPKerberosAuth(mutual_authentication=OPTIONAL),verify=False,allow_redirects=True,timeout=30).json()
    except:
        print("\nEnsure SIM website is opened and accessed in your browser")
        print("\nOnce done, rerun program")
        pass
    
def initialize():
    global today
    today = datetime.datetime.now().strftime("%Y-%m-%d")    
    print("\nGetting SIMs information...\n")
    tasklistlocation = r'\\ant\dept-eu\TBA\UK\Business Analyses\CentralOPS\EU CO Systems\nnmaherr'
    os.chdir(tasklistlocation)
    trialDSfile = csv.reader(open("logins.csv","r"))
    #getting the DS in the file
    schedulernamelist = []
    for lines in trialDSfile:
        schedulernamelist.append(lines[0])




    #PickAndFill
    numberofSIMsresolved=[]
    urlPickandFillclosed = 'https://issues.amazon.com/issues/search?q=containingFolder%3A(23a0ce47-7bb8-421d-8db3-cdcfd8c36e90)+status%3A(Resolved)+lastUpdatedDate%3A(%5BNOW-16HOURS..NOW%5D)&sort=lastUpdatedDate+desc&selectedDocument=71d887b3-6485-4176-b4a1-78bd8ef0b113'
    count = 0
    for schedulername in schedulernamelist:
        url = 'https://maxis-service-prod-pdx.amazon.com/issues?q=containingFolder:(23a0ce47-7bb8-421d-8db3-cdcfd8c36e90)%20assignee:(' +schedulername+ ')%20status:(Resolved)%20lastUpdatedDate:[NOW-16HOURS%20TO%20NOW]&sort=lastUpdatedConversationDate%20desc&selectedDocument=1f923d52-9dc4-4351-a733-44d5ede9d152'
        response = requests.get(url, auth=HTTPKerberosAuth(mutual_authentication=OPTIONAL),verify=False,allow_redirects=True,timeout=30).json()
        numberofSIMsresolved.append((response['totalNumberFound']))
        count += 1
        print("Retrieving information Pick and Fill SIMs Resolved: " + str(((int(count)/int(len(schedulernamelist)))*100))[0:5] + "%", end="\r")
    print("-------------Success: Pick and Fill SIMs Resolved retrieved--------------")

    #PickAndFill
    numberofSIMsopen=[]
    urlPickandFillopen = 'https://issues.amazon.com/issues/search?q=containingFolder%3A(23a0ce47-7bb8-421d-8db3-cdcfd8c36e90)+status%3A(Open)+lastUpdatedDate%3A(%5BNOW-16HOURS..NOW%5D)&sort=lastUpdatedDate+desc&selectedDocument=71d887b3-6485-4176-b4a1-78bd8ef0b113'
    count1 = 0
    for schedulername in schedulernamelist:
        url = 'https://maxis-service-prod-pdx.amazon.com/issues?q=containingFolder:(23a0ce47-7bb8-421d-8db3-cdcfd8c36e90)%20assignee:(' +schedulername+ ')%20status:(Resolved)%20lastUpdatedDate:[NOW-16HOURS%20TO%20NOW]&sort=lastUpdatedConversationDate%20desc&selectedDocument=1f923d52-9dc4-4351-a733-44d5ede9d152'
        response = requests.get(url, auth=HTTPKerberosAuth(mutual_authentication=OPTIONAL),verify=False,allow_redirects=True,timeout=30).json()
        numberofSIMsopen.append((response['totalNumberFound']))
        count1 += 1
        print("Retrieving information Pick and Fill SIMs Opened: " + str(((int(count1)/int(len(schedulernamelist)))*100))[0:5] + "%", end="\r")
    print("-------------Success: Pick and Fill SIMs Opened retrieved--------------")

    


    ##formatting message
    datatable = ""
    for scheduler,numberofSIMsresolved,numberofSIMsopen in zip(schedulernamelist, numberofSIMsresolved,numberofSIMsopen):
        if numberofSIMsresolved!=0:
            datatable += "\n|" + str(scheduler) + "|" + str(numberofSIMsresolved)+ "|" + str(numberofSIMsopen) + "|" + str(numberofSIMsopen+numberofSIMsresolved) + "|"
    
    messageprint = "\nSending webhooks to Pick and Fill Chime Room..."
    
    data = {"Content": "/md \n**Number of SIMs per scheduler Pick and Fill:**\n\n|Scheduler|[Count SIMs Resolved]("+ str(urlPickandFillclosed) + ")|[Count SIMs Opened]("+str(urlPickandFillopen) + ")|Total SIMs|\n|---|---|---|"+ datatable }
    
    sendwebhook(data,messageprint) #sending webhook  

            

    

    




       
    

def sendwebhook(data,messageprint):
    
    result = False
    try:
        print(messageprint)
        urlchimeroom = "https://hooks.chime.aws/incomingwebhooks/b07825bd-0337-42da-b159-8d02e555cb80?token=UVRET3Ftc2Z8MXw0SVY3dzdzMzFsX0RGMUVSREJTWnA2TGh2SDVrRVJHOWhsekNZZU5aUHo4"
        result = False
        session = requests.session()
        params = {'format': 'application/json'}
        response = session.post(urlchimeroom, params=params, json=data)
        if response.status_code == 200:
            result = True

        print("\nWebhooks sent\n")
        return result
        
    except Exception as e:
        print("\nFailed to send Chime message: ", e)
        return result
    
    
    







    

if __name__ == "__main__":
    testbrowser()  
    initialize()

    

