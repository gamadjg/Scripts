
"""
-----layout of script-----
X-Read emails from a csv list
    X-if csv file not selected, close app
X-request input for subject and body of email
    X-If no subject/body entered, will set parameters to empty default values
X-Check if outlook is Open
        X-if not open, open Outlook
"""
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

def Emailer(recipient, subject="", body=""): # default arguments for no errors
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = body
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
    try:
        with open(filepath, encoding='utf-8-sig') as csvfile:
            list = csvfile.readlines()
            for entry in list:
                # If the first line is a header, ignore it
                if re.search(r"email.|Email.", entry):
                    pass
                elif re.search(r"^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$", entry):
                    newln = re.sub(r"\n",'; ', entry)
                    newlist +=newln
    except:
        print('File could not be opened, exiting application')
        exit(1)
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
    os.startfile("outlook")
Emailer(recipients, subject, body)
#------------------Test fields---------------------------------
#filepath = get_list_path()
#subject = 'Test subject'
#body = "Test body"
#user_list = import_list(filepath)
#print(user_list)
#Emailer(body, subject, user_list)
#--------------------------------------------------------------
