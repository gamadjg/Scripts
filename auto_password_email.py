import win32com.client
import datetime as dt
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
#if outlook.GetDefaultFolder(6) == "Inbox":
inbox = outlook.GetDefaultFolder(6)

def find_password_exp_folder(inbox):
    for i in range(inbox.Folders.Count):
        # Print the name of each folder within default email account/inbox
        #print(inbox.Folders[i].name)

        # Find the password expiration folder
        if inbox.Folders[i].name == 'Password Expirations':
            password_exp_folder = i
    return password_exp_folder

#def find_latest_email():

pwd_folder = find_password_exp_folder(inbox)
#print(inbox.Folders[pwd_folder].Items)
messages = inbox.Folders[pwd_folder].Items
last_email = messages.GetLast()
print(last_email.body)
#messages = inbox.Items
#message = messages.GetLast()
#body_content = message.body

#print (body_content)
# Find the password expiration emails

# Read the email and extract the email accounts
    # if no email exists, exit script


# structure email response

# Send email out
