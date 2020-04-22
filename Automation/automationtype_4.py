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
import xlsxwriter, xlrd
import datetime
import openpyxl

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

#sql code to get the required information from the overall data 
sql_ATTACH ='''

SQL CODE 1

'''
sql_SNAP ='''

SQL CODE 2

'''

df_ATTACH = pd.read_sql(sql_ATTACH,cnxn) #reading the sql code through function to save the data in a pandas dataframe
df_SNAP = pd.read_sql(sql_SNAP,cnxn)  

#Below, editing the date columnsto show mm/dd/yy format 
df_ATTACH['Next Inspection Date'] = pd.to_datetime(df_ATTACH['Next Inspection Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_ATTACH['Next Overdue Date'] = pd.to_datetime(df_ATTACH['Next Overdue Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_ATTACH['Start Date'] = pd.to_datetime(df_ATTACH['Start Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_ATTACH['End Date'] = pd.to_datetime(df_ATTACH['End Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_ATTACH.index = np.arange(1,len(df_ATTACH)+1) #to make the dataframe start with 1 instead of 0
df_ATTACH = df_ATTACH[['Inspection Type','County','Inspection Name','End Date','Next Inspection Date','Next Overdue Date','Percent Due']]


df_SNAP['End Date'] = pd.to_datetime(df_SNAP['End Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_SNAP['Next Inspection Date'] = pd.to_datetime(df_SNAP['Next Inspection Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_SNAP['Next Overdue Date'] = pd.to_datetime(df_SNAP['Next Overdue Date'], dayfirst = False, yearfirst = False).dt.strftime('%m/%d/%Y')
df_SNAP = df_SNAP[df_SNAP['Percent Due']>89.99]
df_SNAP.index = np.arange(1,len(df_SNAP)+1) #to make the dataframe start with 1 instead of 0
df_FinalSNAP = df_SNAP[['Parameter1','Parameter2', 'Parameter3']]


writer = pd.ExcelWriter(r'PATH\Report 2 - WM test.xlsx', engine='xlsxwriter')
df_ATTACH.to_excel(writer)
workbook  = writer.book
worksheet = writer.sheets['Sheet1']
book = xlrd.open_workbook(r'PATH\Report 2 - WM test.xlsx')
sheet = book.sheet_by_name('Sheet1')

worksheet.insert_image('J3',r'PATH\WMreport2Image.PNG' )

worksheet.freeze_panes(1,0)

for index, col in enumerate(df_ATTACH):
    mlength = df_ATTACH[col].astype(str).map(len).max()
    new_length = max((
        mlength,  # len of largest item
        len(str(df_ATTACH[col].name))  # len of column name/header
        )) + 2     
    worksheet.set_column(index + 1, index + 1, new_length)



i = 1       #code block for the derived formula and color coding in the attached excel 
j = 4     
#while i != sheet.nrows:
format1 = workbook.add_format({'bold': False, 'font_color': 'black','bg_color': '#74ec65'})
format2 = workbook.add_format({'bold': False, 'font_color': 'black','bg_color': 'yellow'})    
format3 = workbook.add_format({'bold': False, 'font_color': 'black','bg_color': '#fd5353'})
worksheet.conditional_format(i,j+1,sheet.nrows - 1,j+1,{'type': 'formula','criteria':'=$H2<90','format': format1})        
worksheet.conditional_format(i,j+1,sheet.nrows - 1,j+1,{'type': 'formula','criteria':'=OR($H2=90,$H2=100,And($H2>90,$H2<100))','format': format2})
worksheet.conditional_format(i,j+1,sheet.nrows - 1,j+1,{'type': 'formula','criteria':'=OR($H2=101,$H2=125,And($H2>101,$H2<125))','format': format3})
writer.save()

#to_list = (['Tapasya.sharma'])
UID = getpass.getuser() + '@gmail.com'

print("------ To list from .txt file -----")
txt_list = open(r'PATH\to_list.txt', 'r').readlines()
to_list = []
for i in txt_list:
    to_list.append(i.strip('\n'))
    print(i.strip('\n'))
print("------ To list from .txt file -----")


#print (UID)
IA = df_ATTACH.shape[0]

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
df_Q = pd.read_excel(r'PATH//RandomQuotes.xlsx',sheet_name ='Sheet1')
df_Q = df_Q[['Quote']]
df_Q.set_index('Quote', inplace=True)

        
def send_email(TO_LIST, SUBJECT = None, MESSAGE = None):
    server = smtp_connect('smtp.office365.com')
    if server is None:
        print("Script cancelled, have a good day, kind sir!")
        sys.exit()
        
    if MESSAGE is None:
        MESSAGE = 'Found ' +str(IA) +'relevant WO to job, Report Ran at ' + str(datetime.datetime.now().strftime("%H:%M:%S")) + " on " + str(datetime.datetime.now().strftime("%m/%d/%y") + '.')
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
            if i == 'John.smith' or i == 'tapasya.sharma':
                i = i + '@gmail.com'
            else:
                i = i + '@yahoo.com'
    i_list.append(i)
    i_list_str = i_list_str + i
    m = MIMEMultipart()
            
    html = """\
    <html>
        <head>
        <span style = "color:Tomato;"><b>WARNING</span> - This email/attachments may contain non-public information.</b>
        <body>
        <br>
        <h4><span style = "color:#00008b;">The purpose  of this email is to provide snapshots of the Demarcated data by percentages of FORWARD LOOKING DASHBOARD - UNIT SUBSTATION </span> </h4> 
        Python has Scanned Database and is now emailing you. <br>
        The following list is sorted by the percent due and are all 100% and above and fall FORWARD LOOKING DASHBOARD.</br>
        <br>
        Sincerly,<br>
        PECO Python
        <br>
        <br>
        {0}
        <br>                
        <center>Quote of the day!</center>
        <br>
        <center>{1}</center>


        </body>
    <html>
    """.format(df_FinalSNAP.to_html(),df_Q.sample().to_html())  #df_SNAP.sample().to_html()

            
    m['Subject'] = 'Demarcated data by percentages of FORWARD LOOKING DASHBOARD - UNIT SUBSTATION   - ' +str(IA) +' Inspection found on ' + str(datetime.date.today().strftime("%m/%d/%y"))
    m['From'] = UID
    m['To'] = ', '.join(i_list)

    part = MIMEBase('application', "octet-stream")
    part1 = text(html,'html')
    part.set_payload(open(r'PATH\Report 2 - WM test.xlsx', "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="Report.xlsx"')
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
        send_email(to_list)
        print("Emails sent, enjoy your day :)")
        #sys.exit()
    else:
        print('Very well, emails not sent.')
       #sys.exit()
else:
    print("No errors found.")



