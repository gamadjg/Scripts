import win32com.client
import datetime as dt

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

def get_expiring_emails_from_body(body):
    list = body.split()
    print(list)
    return ''

#----------------------------RUNNING----------------------------
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
#if outlook.GetDefaultFolder(6) == "Inbox":
inbox = outlook.GetDefaultFolder(6)
pwd_folder = find_password_exp_folder(inbox)
body = find_latest_email(pwd_folder)
exp_emails = get_expiring_emails_from_body(body)

#----------------------------TESTING----------------------------
# Print the name of each folder within default email account/inbox
#for i in range(inbox.Folders.Count):
    #print(inbox.Folders[i].name)
