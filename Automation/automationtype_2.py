
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

# Basic cascade ODBC connection info
serv = '
host = '
port = 
db_name = 
user_id = 
pwd = 

dsn_name = 
cnxn = pyodbc.connect('DSN=' + dsn_name + ';'
                      'UID=' + user_id + ';'
                      'PWD=' + pwd)




sql_NUC= '''
SQL CODE
''' 
   

# Read into a Namplate Pandas Dataframe
df_NUC = pd.read_sql(sql_NUC,cnxn)


df_NUC['Date'] = pd.to_datetime(df_NUC['Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_NUC['Schedule'] = pd.to_datetime(df_NUC['Schedule'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_NUC['last_dt'] = pd.to_datetime(df_NUC['last_dt'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_NUC['Overdue_Date'] = pd.to_datetime(df_NUC['Overdue_Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')


df_NUC = df_NUC[(df_NUC['Sch_Percent'].isnull())| (df_NUC['Sch_Percent']>112.5)]
df_NUC = df_NUC[df_NUC['Today_Percent']>93]
df_NUC.sort_values(['Sch_Percent'], axis=0, ascending=False, inplace=True)


IA = df_Final.shape[0]


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
            if i == 'Tapasya.Sharma' or i == 'John.Smith':
                i = i + '@yahoo.com'
            else:
                i = i + '@gmail.com'
            
            m = MIMEMultipart()
            html = """\
            <html>
                <head>Python has Scanned database and is now emailing you.<br> 
                <p style="color:#00008b;">Team,</p>

                <p style="color:#00008b;">The Dashboard was run and the following items are scheduled in grace.  Moving the work order could Jeopardize compliance </p>
                <br>
                <body>
                 {0}            
                 <br>
                 <br>
                 Sincerly,<br>
                 PECO Python
                </body>
            </html>
            """.format(df_Final.to_html())
            part1 = text(html,'html')
            m.attach(part1)
            m['Subject'] = 'DASHBOARD STATUS NOTIFICATION ' +str(IA) +' Found'
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
    print("No errors found.")
    
    


