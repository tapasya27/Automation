
#v8 12/11/19 - 10:42 AM 

'''
Programmers: Tapasya Sharma
Date: 12/10/2019
Purpose:  
'''


from smtplib import *
import smtplib,ssl
from email.mime.text import MIMEText as text
from email.mime.application import MIMEApplication as application
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
import smtplib, sys, datetime, os, getpass
import pyodbc, pandas as pd, numpy as np
import io 
from email import encoders
from email.utils import formatdate 
import random 
import datetime

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

sql_PUC ='''
SQL CODE 1'''

sql_PUC2 ='''
SQL CODE 2'''

df_PUC = pd.read_sql(sql_PUC,cnxn)
df_PUC2 = pd.read_sql(sql_PUC2,cnxn)

df_PUC['Last Date'] = pd.to_datetime(df_PUC['Last Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_PUC['Schedule Date'] = pd.to_datetime(df_PUC['Schedule Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_PUC['OrigDueDate'] = pd.to_datetime(df_PUC['OrigDueDate'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_PUC['OrigOverDueDate'] = pd.to_datetime(df_PUC['OrigOverDueDate'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')

df_PUC2['Last Date'] = pd.to_datetime(df_PUC2['Last Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_PUC2['Schedule Date'] = pd.to_datetime(df_PUC2['Schedule Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_PUC2['OrigDueDate'] = pd.to_datetime(df_PUC2['OrigDueDate'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_PUC2['OrigOverDueDate'] = pd.to_datetime(df_PUC2['OrigOverDueDate'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')


df_Attach = df_PUC[['Parameter1','Parameter2', 'Parameter3']]
df_Attach2 = df_PUC2[['Parameter1','Parameter2', 'Parameter3']]

def highlight_vals(val):
    color='yellow'
    if  val > 99:
        return 'background-color: %s' % color
    else:
        return ''

html_table = (
    df_Attach2.style.applymap(highlight_vals, subset=['Percent Due']).set_table_attributes("border=1").render())


df_Attach.index = np.arange(1,len(df_Attach)+1)
df_Attach2.index = np.arange(1,len(df_Attach2)+1)
writer = pd.ExcelWriter(r'PATH/sample.xlsx', engine='xlsxwriter')
df_Attach.to_excel(writer)
workbook = writer.book
#print("a")
#print("b")
worksheet = writer.sheets['Sheet1'] # pull worksheet object
#print("c")
worksheet.freeze_panes(1,0)
worksheet.set_landscape()
worksheet.set_zoom(90)
worksheet.fit_to_pages(1, 1)

for idx, col in enumerate(df_Attach):  # loop through all columns
    series = df_Attach[col]
    l = max((
        series.astype(str).map(len).max(),  # len of largest item
        len(str(series.name))  # len of column name/header
        ))
    if l > 17:
        delta = 8
    else: 
        delta = 1
        
    max_len = max((
        series.astype(str).map(len).max(),  # len of largest item
        len(str(series.name))  # len of column name/header
        )) + delta  # adding a little extra space
    worksheet.set_column(idx+1, idx+1, max_len)  # set column width
header_format = workbook.add_format({
    'bold': True,
    'text_wrap': True,
    'bg_color': 'yellow',
    'border': 1})


for col_num, value in enumerate(df_Attach.columns.values):
    worksheet.write(0, col_num + 1, value, header_format)

writer.save()
#print("d")

#print("e")

to_list = (['Tapasya.Sharma'])

UID = getpass.getuser() + '@gmail.com'

IA = df_Attach2.shape[0]

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
df_Q = pd.read_excel('PATH/RandomQuotes.xlsx',sheet_name ='Sheet1')
df_Q = df_Q[['Quote']]
df_Q.set_index('Quote', inplace=True)

        
def send_email(TO_LIST, SUBJECT = None, MESSAGE = None):
    server = smtp_connect('smtp.office365.com')
    if server is None:
        print("Script cancelled, have a good day, kind sir!")
        sys.exit()
        
    if MESSAGE is None:
        MESSAGE = 'Found ' +str(IA) +' Report Ran at ' + str(datetime.datetime.now().strftime("%H:%M:%S")) + " on " + str(datetime.datetime.now().strftime("%m/%d/%y") + '.')
    else:
        MESSAGE = MESSAGE
    if SUBJECT is None:
        SUBJECT = 'Exercise Mechanism Moves'
    else:
        SUBJECT = SUBJECT
    i_list = []
    i_list_str = ""
    for i in TO_LIST:
        
        i = str(i)
        if i is not None and "nan" not in i.lower() and i.lower() is not "":
            if i == 'John.Smith' or i=='Tapasya.sharma':
                i = i + '@gmail.com,'
            else:
                i = i + '@yahoo.com,'
        i_list.append(i)
        i_list_str = i_list_str + i
                
    m = MIMEMultipart()

    html = """\
    <html>
        <head>
        <span style = "color:Tomato;"><b>WARNING</span> - This email/attachments may contain non-public information Warning!</b>
        <body>
        <br>
        <h4><span style = "color:#00008b;">The purpose  of this email is to provide snapshots of process.</span> </h4> 
        Python has Scanned Database and is now emailing you. <br>
        The following list is sorted by the percent due and are all 90%.</br>
        <br>
        Sincerly,<br>
        XX
        <br>
        <br>
        {0}
        <br>                
        <center>Quote of the day!</center>
        <br>
        <center>{1}</center>


        </body>
    <html>
    """.format(html_table,df_Q.sample().to_html())


    m['Subject'] = 'Dashboard inspections - ' +str(IA) +' Workorders found on ' + str(datetime.date.today().strftime("%m/%d/%y"))
    m['From'] = UID
    m['To'] = ", ".join(i_list)

    part = MIMEBase('application', "octet-stream")
    part1 = text(html,'html')
    part.set_payload(open(r'Path/sample.xlsx', "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="sample.xlsx"')
    m.attach(part)
    m.attach(part1)
    server.sendmail(UID, i_list, m.as_string()) 
    print('sent to: ' + i_list_str)
    print('---------------------')
    
    server.quit()
    

if IA > 0:
    yn = 'y'
else:
    yn = 'n'

if 'y' in yn:
    permission = input('Emails to be sent from ' + UID + ' to the above list, is that correct (y/n)? ')
    if 'y' in permission.lower():
        print("Enter your Passcode:")
        send_email(to_list)
        print("Emails sent, enjoy your day :)")
        #sys.exit()
    else:
        print('Very well, emails not sent.')
       #sys.exit()
else:
    print("No errors found.")
#df_final.head()
#df.dtypes




