"""
-----Outline-----
Finds the defaul Outlook account
    if not admin email, close script
    look for admin email
Looks at folders within account and finds password expiration folder
    if password expiration folder does not exist, close script
Finds latest password expiration email
    if it doesnt exist, notify and close
    if no expiring users, notify and close
temp: auto opens outlook and creates emails to each sender for review
    preferred: auto send emails to expiring users
"""

import win32com.client
import win32ui
import datetime as dt
import re
#----------------------------Variable Declaration------------------------
OKTA_URL = "https://alkahest.okta.com/"
OKTA_URL_TEXT = "alkahest.okta.com"
SENDER_FIRST_NAME ="Alkahest Admin"
EMAIL_BODY = """\
<html>
  <head></head>
  <body>
    <p>
       This is an automatic email to notify you that your password will be expiring
       in {} days.<br>

       Please reset your password by:<br>
       <ol>
        <li>Searching for <a href="{}">{}</a> within your preferred browser.</li>
        <li>Loging in with your email address and expiring password.</li>
            <ul><li>If your password has already expired, you will immediately be prompted to reset your password upon login.</li></ul>
        <li>Once logged in, on the top right of the screen click on your <i>First name/Settings</i></li>
        <li>Here you will see a box labaled <i>Change Password</i> which requires you to enter your current and new password.</li>
       </ol>

       Thanks,<br>
       {}
    </p>
  </body>
</html>
"""
#----------------------------Functions----------------------------
def outlook_is_running():
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
        return True
    except win32ui.error:
        return False

def find_password_exp_folder(inbox):
    for i in range(inbox.Folders.Count):
        if inbox.Folders[i].name == 'Password Expirations':
            password_exp_folder = i
    return password_exp_folder

def find_latest_email(password_exp_folder):
    messages = inbox.Folders[pwd_folder].Items
    last_email = messages.GetLast()
    body = last_email.body
    return body

def get_expiring_accounts_from_body(exp_pwd_email):
    #print(exp_pwd_email)
    results = re.findall(r"([a-z.]*@\w*.com);\t(\d)",exp_pwd_email)
    #print(results)
    return results

def compose_emails(recipients):
    outlook = win32com.client.Dispatch('outlook.application')
    print(len(recipients))
    for recipient in range(len(recipients)):
        mail = outlook.CreateItem(0)
        mail.To = recipients[recipient][0]
        mail.Subject = 'Auto-Notification: Password Expiration'
        mail.HtmlBody = EMAIL_BODY.format(recipients[recipient][1], OKTA_URL, OKTA_URL_TEXT, SENDER_FIRST_NAME)
        mail.Display(True)
    exit()

#----------------------------RUNNING----------------------------
if not outlook_is_running():
    import os
    os.startfile("outlook")
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
pwd_folder = find_password_exp_folder(inbox)
body = find_latest_email(pwd_folder)
exp_emails = get_expiring_accounts_from_body(body)
compose_emails(exp_emails)
#----------------------------TESTING----------------------------
# Print the name of each folder within default email account/inbox
#for i in range(inbox.Folders.Count):
    #print(inbox.Folders[i].name)
