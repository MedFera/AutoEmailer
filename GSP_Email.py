import pandas as pd
import datetime as dt
from datetime import timedelta
import yagmail

class Email:
    def __init__(self):
        #Enter Email and alternate access password here
        self.yag = yagmail.SMTP('email','password')
        self.receivers = []
        self.body = ""
        self.files = []
        self.title = "Subject Head"
        
    def sendMail(self):
        self.yag.send(
            to = self.receivers,
            subject = self.title,
            contents = self.body,
            attachments = self.files,
        )

file_path = 'Assigned eCards GSPGuard.xlsx'
sheets = ['ACLS Course', 'BLS Course', 'FA.CPR.AED', 'CPR.AED Only', 'Lorain Cards',]
for y in sheets:
    sheetName = y
    dataframe = pd.read_excel(file_path,sheet_name=sheetName)

    #cd = column data
    date_cd = dataframe['Course Date']
    email_cd = dataframe['Email']
    fname_cd = dataframe['First Name']
    lname_cd = dataframe['Last Name']

    #combined all column data
    totalData = pd.concat([date_cd, fname_cd, email_cd, lname_cd],axis=1)

    #print(totalData)

    #checks 715 days prior 
    expireDate = dt.datetime.now() - timedelta(days=715)

    #creates an array of index values of licenses that expired
    expiredLicenses = totalData.index[totalData['Course Date'] < expireDate ]

    #PRINT CHECKS
    #print(expiredLicences)

    #print(expireDate)

    #'gspguard@gmail.com','nqoaqwkrdhppdajx'


    

    #Goes through Expired Licenc
    for x in expiredLicenses :
        date = totalData.loc[x,'Course Date']
        fname = totalData.loc[x,'First Name']
        lname = totalData.loc[x, 'Last Name']
        email = totalData.loc[x,'Email']
        fullName = fname + " " + lname
        
        sentence = "Hello " + fullName + ". Your "+ sheetName +" license(s) expires " + date.strftime("%B %d, %Y")+". Please contact me to renew you license!"
        
        newEmail = Email()
        newEmail.receivers.append(email)
        newEmail.body = sentence
        newEmail.title = sheetName + " Renewal" 
        newEmail.sendMail()
        print("Emailed "+ fullName +" at: "+ email + " for " + sheetName + " Renewal" )
        
    #Send Email to Gezim    
    if (sheetName == 'ACLS Course' or sheetName == 'BLS Course'):
        sentence = sheetName + "\n"
        for x in expiredLicenses :
            date = totalData.loc[x,'Course Date']
            fname = totalData.loc[x,'First Name']
            lname = totalData.loc[x, 'Last Name']
            email = totalData.loc[x,'Email']
            fullName = fname + " " + lname
            sentence += fullName + " expired! Last Course Date: " + date.strftime("%B %d, %Y") + "\n"
        print(sentence)
        
        newEmail2 = Email()
        newEmail2.receivers.append('gspguard@gmail.com')
        newEmail2.body = sentence
        newEmail2.title = dt.datetime.now().strftime("%B %d, %Y")+" - Expired " + sheetName 
        newEmail2.sendMail()
    