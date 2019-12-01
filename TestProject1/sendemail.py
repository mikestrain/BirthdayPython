# Import smtplib for the actual sending function
import os
import smtplib
import numpy
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import openpyxl

os.system('cls')

wb2 = openpyxl.load_workbook("Birthday Reminder List.xlsx", data_only = True)
sheet = wb2.get_sheet_by_name('Sheet1')

for i in range(0, 40):
    ii = i + 10

    due = sheet['F' + str(ii)].value

    if due == 1:
        person = sheet['B' + str(ii)].value
        event = sheet['C' + str(ii)].value
        date = sheet['D' + str(ii)].value
        
try:
    
    EMAIL_ADDRESS = 'mike.birthdayreminder@gmail.com'
    EMAIL_PASSWORD = 'Birthday123!'

    server = smtplib.SMTP('smtp.gmail.com', 587)

    server.ehlo()
    server.starttls()
    server.login(EMAIL_ADDRESS,EMAIL_PASSWORD)

    subject = 'Birthday Reminder for' + person
    msg = 'Hello Mike, \nPlease remember the following birthday: \n\nPerson: ' + person + '\nEvent: ' + event + '\nDate:' + str(date)[5:10] + '\n\nThis message has been sent by a reminder robot!'
    print(msg)
    message = 'Subject: {}\n\n{}'.format(subject,msg)

    server.sendmail(EMAIL_ADDRESS, ['mstrain@busek.com', EMAIL_ADDRESS], message)
    server.quit()
    print('\nsuccess!')

except:
    print('email failed to send')
