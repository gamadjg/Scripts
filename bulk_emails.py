# To send out bulk emails to all staff
# Read emails from a csv list
# Create body of email & add attachments if necessary
# Can we open outlook and auto-insert all emails from the list as people to send to?
import win32com.client as win32
import win32ui
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import csv
import re
import pandas as pd

def outlook_is_running():
    try:
        win32ui.FindWindow(None, "Microsoft Outlook")
        return True
    except win32ui.error:
        return False

def Emailer(text, subject, recipient):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    mail.Display(True)
    exit()

def get_list_path():
    # Open file browser GUI to select email list
    Tk().withdraw() # Prevent full GUI from appearing
    filepath = askopenfilename() # Return selected file path
    return filepath

def import_list(filepath):
    list = ''
    newlist=''
    with open(filepath, encoding='utf-8-sig') as csvfile:
        list = csvfile.readlines()
        for entry in list:
            # If the first line is a header, ignore it
            if re.search(r"email.|Email.", entry):
                pass
            elif re.search(r"^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$", entry):
                newln = re.sub(r"\n",'; ', entry)
                newlist +=newln
    # ------------Test with pandas-----------------------------
    #df = pd.read_csv(filepath)
    #email_col = df['Emails']
    #return email_col.values.tolist()
    return newlist

#-------------------Working-------------------------------------

filepath = get_list_path()
recipients = import_list(filepath) # Import list needs to be one column with no headers
subject = input("Enter subject: ")
body = input("Enter base body: ")
if not outlook_is_running():
    import os
    os.startfile("outlook")
Emailer(body, subject, recipients)
#------------------Test fields---------------------------------
filepath = get_list_path()
#subject = 'Test subject'
#body = "Test body"
#user_list = import_list(filepath)
#print(user_list)
#Emailer(body, subject, user_list)
#--------------------------------------------------------------
