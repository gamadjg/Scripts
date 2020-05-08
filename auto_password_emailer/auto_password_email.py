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
from datetime import date
import re
import time
import os
from auto_password_emailer import site
# ----------------------------Functions----------------------------


def outlook_is_running():
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
    except win32ui.error:
        os.startfile("outlook")
        time.sleep(4)


def find_admin_inbox():
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts = win32com.client.Dispatch("Outlook.Application").Session.Accounts
    for account in accounts:
        if account.DisplayName == site.get_email_address():
            account_folders = outlook.Folders(account.DeliveryStore.DisplayName)
            folders = account_folders.Folders
            for folder in folders:
                if folder.name == 'Inbox':
                    inbox = folder
                    return inbox
        else:
            print('Admin account does not exist, exiting application.')
            exit()


def find_password_exp_folder(inbox):
    for i in range(inbox.Folders.Count):
        if inbox.Folders[i].name == 'Password Expirations':
            password_exp_folder = i
            return password_exp_folder
    # If the Password Expirations folder does not exist
    print('Password Expirations folder does not exist. Exiting application.')
    exit(1)


def find_latest_email(password_exp_folder):
    last_email = inbox.Folders[pwd_folder].Items.GetLast()
    last_email_date = re.search(r"\d\d\d\d-\d\d-\d\d", str(last_email.SentOn))
    if last_email_date.group(0) == date.today():
        body = last_email.body
        return body
    else:
        print('No email from today. Exiting application.')
        exit(1)


def get_expiring_accounts_from_body(exp_pwd_email):
    results = re.findall(r"([a-z.]*@\w*.com);\t(\d)",exp_pwd_email)
    if not results:
        print('No upcoming expiring email accounts. Exiting.')
        exit(1)
    else:
        return results


def compose_emails(recipients):
    outlook = win32com.client.Dispatch('outlook.application')
    for recipient in range(len(recipients)):
        mail = outlook.CreateItem(0)
        mail.To = recipients[recipient][0]
        mail.Subject = 'Auto-Notification: Password Expiration'
        mail.HtmlBody = site.get_email_body().format(recipients[recipient][1],
                                                     site.get_okta_url(),
                                                     site.get_okta_text(),
                                                     site.get_sender()
                                                     )
        mail.Display(True)
    print('pass')
    exit()


# ----------------------------RUNNING----------------------------
site = site.SiteInfo()
outlook_is_running()
inbox = find_admin_inbox()
pwd_folder = find_password_exp_folder(inbox)
body = find_latest_email(pwd_folder)
exp_emails = get_expiring_accounts_from_body(body)
compose_emails(exp_emails)