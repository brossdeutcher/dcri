#! python3
##
## emailHandler.py | stores email templates & creates drafts in Outlook

import win32com.client as win32
import pandas as pd
import datetime


def create_draft(text, subject, recipient, cc, attachments=False):

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    #mail.SentOnBehalfOfName = ''	## optional: change the From to a non-default address
    mail.To = recipient
    mail.CC = cc
    mail.Subject = subject
    mail.HtmlBody = text
    if attachments:
        for att in attachments.split(', '):
            print(att)
            mail.Attachments.Add(att)
    mail.Display(True)


dfEmail = pd.read_csv(r'N:\\')	## string replaced with CSV file
print(dfEmail)  # report data displayed in console

def createMail(reportNum: int):
    today = datetime.date.today()
    curMDY = today.strftime("%d%m%Y")
    curDay = today.strftime("%d")
    curMonInt = today.strftime("%m")
    curMonLet = today.strftime("%b")
    curYr = today.strftime("%Y")
    print('\nCreating draft for ' + str(dfEmail.loc[reportNum][0]) + '...')    # report name displayed to console
    print(today)  # date displayed to console
    dateList = []
    for datePt in str(dfEmail.loc[reportNum][6]).split(', '):
        dateList.append(eval(datePt))
    create_draft(
        dfEmail.loc[reportNum][4]   # body
        ,dfEmail.loc[reportNum][3]  # subject
        ,dfEmail.loc[reportNum][1]  # to
        ,dfEmail.loc[reportNum][2]  # cc
        ,dfEmail.loc[reportNum][5] % tuple(dateList)  # attachments
        )
    print('createMail() to rerun')


createMail(int(input('\nEnter report number: ')))

