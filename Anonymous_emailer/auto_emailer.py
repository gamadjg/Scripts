"""
-----Outline-----
X-Checks is Outlook is open, if not then open
X-Finds the defaul Outlook account
    X-look for admin email, if id doesn't exist, close app
X-Looks at folders within account and finds password expiration folder
    X-if password expiration folder does not exist, close script
X-Finds latest password expiration email
    X-if it doesnt exist, notify and close
X-Pull users and expiration dates
    X-if no users, notify and exit.
X-current: auto opens outlook and creates emails to each sender for review
    X-preferred: auto send emails to expiring users
"""

# ----------------------------Importing----------------------------
import win32com.client
import win32ui
import datetime
import pandas
import re
import time
import os
import site_info
# ----------------------------Functions----------------------------


def outlook_is_running():
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
    except win32ui.error:
        os.startfile("outlook")
        time.sleep(4)


def find_inbox():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts
    for account in accounts:
        if account.DisplayName == site.get_email_address():
            account_folders = outlook.Folders(account.DeliveryStore.DisplayName)
            account_folders = account_folders.Folders
            for folder in account_folders:
                if folder.name == 'Inbox':
                    return folder
    else:
        print('Inbox account does not exist, exiting application.')
        exit()

def find_todays_emails(inbox):
    messages = inbox.Items
    msg = messages.GetLast()
    msg_datetime = msg.SentOn
    msg_date = re.search(r"\d\d\d\d-\d\d-\d\d", str(msg_datetime))
    today = str(datetime.date.today())
    subject = []
    body = []
    cont = True
    if msg_date.group(0) != today:
        # Do nothing, no messages today.
        return 0, 0
    else:
        while cont:
            subject.append(msg.subject+site.get_domain())
            body.append(re.sub(r'\r\n\r\n', r'\r\n', msg.body))
            # Check to see if there are any more messages
            msg = messages.GetPrevious()
            if msg is None:
                print('no emails left')
                cont = False
            else:
                print("emails still exist")
                msg_datetime = msg.SentOn
                msg_date = re.search(r"\d\d\d\d-\d\d-\d\d", str(msg_datetime))

                if msg_date.group(0) != today:
                    cont = False
        return subject, body


def create_df(sbjct, bdy):
    df = pandas.DataFrame()
    df['Subject'] = sbjct
    df['Body'] = bdy
    return df

# ----------------------------RUNNING----------------------------
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
site = site_info.SiteInfo()
outlook_is_running()
inbox = find_inbox()
subject, body = find_todays_emails(inbox)
if type(subject) is list:
    outlook = win32com.client.Dispatch('outlook.application')
    df = create_df(subject, body)
    for i in range(len(df)):
        mail = outlook.CreateItem(0)
        mail.To = df['Subject'].iloc[i]
        mail.Subject = "You've got mail! Anonymous Recognition!"
        mail.Body = df['Body'].iloc[i]
        # Display email in Outlook, manually send out
        mail.Display(True)
        # Send out created email without looking at it
        # mail.Send()
else:
    print("No emails today.")
