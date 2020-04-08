import win32com.client
import datetime as dt

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder

password_exp_folder = ""

#print(inbox.Folders.Count)
#print(inbox.Folders[5].name)

for i in range(inbox.Folders.Count):
    # Get the return a list of all the folders within the main email account/inbox
    #print(inbox.Folders[i].name)

    if inbox.Folders[i].name == 'Password Expirations':
        password_exp_folder = i

print(inbox.Folders[password_exp_folder].Items)


#messages = inbox.Items
#message = messages.GetLast()
#body_content = message.body

#print (body_content)
# Find the password expiration emails

# Read the email and extract the email accounts
    # if no email exists, exit script


# structure email response

# Send email out
