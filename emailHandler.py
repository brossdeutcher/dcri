#! python3
##
## emailHandler.py | stores email templates & creates drafts in Outlook

import win32com.client as win32
import pandas as pd
import datetime


def create_draft(text, subject, recipient, cc, attachments=False):

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.SentOnBehalfOfName = 'DCLIDomesticReporting@dcli.com'
    mail.To = recipient
    mail.CC = cc
    mail.Subject = subject
    mail.HtmlBody = text
    if attachments:
        for att in attachments.split(', '):
            print(att)
            mail.Attachments.Add(att)
    mail.Display(True)


dfEmail = pd.read_csv(r'N:\ ')  ## file path in here
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
    print('createMail() to run a new report!')

##createMail(int(input('\nEnter report number: '))) ## this can run createMail if used in console


## Tkinter User Interface

from tkinter import *
from functools import partial

def sel():
   selection = "You selected: " + dfEmail.loc[var.get()][0]
   label.config(text = selection)
   button.config(command=partial(createMail, int(var.get())))

root = Tk()
root.title('DCLI Report Email Distributions')
var = IntVar()
i = 0

for report in range(len(dfEmail)):
    Radiobutton(root, text=dfEmail.loc[i][0], variable=var, value=i, command=sel).pack(anchor=W)
    i += 1

label = Label(root, text=dfEmail.loc[var.get()][0])
label.pack()

button = Button(master=root, text='Run Selected Report!', command=partial(createMail, int(var.get())))
button.pack()

root.mainloop()

