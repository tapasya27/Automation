'''
Author: Tapasya Sharma
Date: 11/12/2019
Purpose: 


'''


from smtplib import *
from email.mime.text import MIMEText as text
from email.mime.application import MIMEApplication as application
from email.mime.multipart import MIMEMultipart 
import smtplib, sys, datetime, os, getpass
import pyodbc, pandas as pd, numpy as np
import datetime as dt
import time


# Basic cascade ODBC connection info
serv = 
host = 
port = 
db_name = 
user_id = 
pwd = 

dsn_name = 
cnxn = pyodbc.connect('DSN=' + dsn_name + ';'
                      'UID=' + user_id + ';'
                      'PWD=' + pwd)



# Specific SQL statment to extract the battery closure information from Cascade

sql_5w = '''
SQL CODE 1 '''

sql_cap = '''
SQL CODE 2'''

sql_bannual = '''
SQL CODE 3'''

sql_sannual = '''
SQL CODE 4
'''



# Read into a Namplate Pandas Dataframe
df_5w = pd.read_sql(sql_5w,cnxn)
df_cap = pd.read_sql(sql_cap,cnxn)
df_bannual = pd.read_sql(sql_bannual,cnxn)
df_sannual = pd.read_sql(sql_sannual,cnxn)



df_5w['Date'] = pd.to_datetime(df_5w['Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_5w['Closed_Date'] = pd.to_datetime(df_5w['Closed_Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_5w = df_5w[df_5w['Ok/Fix']=="fix"]


df_cap['Date'] = pd.to_datetime(df_cap['Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_cap['Date'] = pd.to_datetime(df_cap['Closed_Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_cap = df_cap[df_cap['Ok/Fix']=="fix"]


df_bannual['Date'] = pd.to_datetime(df_bannual['Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_bannual['Closed_Date'] = pd.to_datetime(df_bannual['Closed_Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_bannual = df_bannual[df_bannual['Ok/Fix']=="fix"]

df_sannual['Date'] = pd.to_datetime(df_sannual['Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_sannual['Closed_Date'] = pd.to_datetime(df_sannual['Closed_Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_sannual = df_sannual[df_sannual['Ok/Fix']=="fix"]


dfT = df_5w.append(df_cap, ignore_index=True)
dfT2 = dfT.append(df_bannual, ignore_index=True)
df_final = dfT2.append(df_sannual, ignore_index=True)


IA = df_final.shape[0]

YY = time.strftime("%y", time.localtime())
WorkWeek = (dt.date.today().strftime("%V"))


#to_list = ([,'Tapasya.Sharma'])    
to_list = (['Tapasya.Sharma'])   
    


UID = getpass.getuser() + '@gmail.com'



msg = MIMEMultipart()


def smtp_connect(SMTP_info):
    server = smtplib.SMTP(SMTP_info)
    server.starttls()

    pwd_allowance = True
    while pwd_allowance is True:
        PWD = getpass.getpass()
        try:
            server.login(UID, PWD)
        except smtplib.SMTPAuthenticationError:
            yn = input('Wrong password! Try again (y/n)? ')
            if 'n' in yn:
                pwd_allowance = False; break #break just in case
                return None
        else:
            pwd_allowance = False
            return server

def send_email(TO_LIST, SUBJECT = None, MESSAGE = None):
    server = smtp_connect('smtp.office365.com')
    if server is None:
        print("Script cancelled, have a good day, kind sir!")
        sys.exit()

    

    for i in TO_LIST:
        i = str(i)
        if i is not None and "nan" not in i.lower() and i.lower() is not "":
            if i == 'John.smith' or i == 'Tapasya.Sharma':
                i = i + '@yahoo.com'
            else:
                i = i + '@gmail.com'
            
            m = MIMEMultipart() 
            
            html = """\
            <html>
                <head>Python has Scanned Database and is now emailing you.<br> 
                <p style="color:red;">REPORT PURPOSE: </p>
                
            
                <body>
                <h3>1) Dataframe 1.</h3>
                 {0}
                <h3>2) Dataframe 2.</h3>
                 {1}
                <h3>3) Dataframe 3.</h3>
                 {2}
                <h3>4) Dataframe 4.</h3>
                 {3}
                 <br>
                 <br>
                 Sincerly,<br>
                Python
                </body>
            </html>
            """.format(df_5w.to_html(),df_cap.to_html(),df_bannual.to_html(),df_sannual.to_html())
            part1 = text(html,'html')
            m.attach(part1)
            m['Subject'] = 'Subject : Issue ' +str(IA) +' Found - WW:'+str(YY)+str(WorkWeek) 
            m['From'] = UID
            m['To'] = i
            server.sendmail(UID, i, m.as_string())
            print('sent to: ' + i)
    print('---------------------')

    server.quit()

#IF NO ERROR FOUND
def noerror_email(TO_LIST, SUBJECT = None, MESSAGE = None):
    server = smtp_connect('smtp.office365.com')
    if server is None:
        print("Script cancelled, have a good day, kind sir!")
        sys.exit()

    

    for i in TO_LIST:
        i = str(i)
        if i is not None and "nan" not in i.lower() and i.lower() is not "":
            if i == 'John.smith' or i == 'Tapasya.Sharma': :
                i = i + '@yahoo.com'
            else:
                i = i + '@google.com'
            
            m = MIMEMultipart() 
            
            html = """\
            <html>
                <head>Python has Scanned Database and is now emailing you.<br> 
                <p style="color:red;">REPORT PURPOSE:</p>
                
            
                <body>
                <h3>1) Dataframe 1 - "NONE"</h3>
                 
                <h3>2) Dataframe 2 - "NONE"</h3>
                 
                <h3>3) Dataframe 3 - "NONE"</h3>
                 
                <h3>4) Dataframe 4 - "NONE"</h3>
                 
                 <br>
                 <br>
                 Sincerly,<br>
                 Python
                </body>
            </html>
            """
            part1 = text(html,'html')
            m.attach(part1)
            m['Subject'] = 'Subject '+ 'WW:'+str(YY)+str(WorkWeek)
            m['From'] = UID
            m['To'] = i
            server.sendmail(UID, i, m.as_string())
            print('sent to: ' + i)
    print('---------------------')

    server.quit()



if IA > 0:
    yn = 'y'
else:
    yn = 'n'

if 'y' in yn:
    permission = input('Emails to be sent from ' + UID + ' to the above list, is that correct (y/n)? ')
    if 'y' in permission.lower():
        send_email(to_list)
        print("Emails sent, enjoy your day :)")
        #sys.exit()
    else:
        print('Very well, emails not sent.')
       #sys.exit()
else:
    permission = input('Emails to be sent from ' + UID + ' to the above list, is that correct (y/n)? ')
    if 'y' in permission.lower():
        noerror_email(to_list)
        print("Emails sent, enjoy your day :)")
        #sys.exit()
    else:
        print('Very well, emails not sent.')
#df_final.head()
#df.dtypes





















