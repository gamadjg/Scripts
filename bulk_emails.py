# To send out bulk emails to all staff

# Read emails from a csv list

# Create body of email & add attachments if necessary

# Can we open outlook and auto-insert all emails from the list as people to send to?
import win32com.client as win32
import os

def Emailer(text, subject, recipient):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    mail.Subject = subject
    mail.HtmlBody = text
    mail.Display(True)
    print('success')
    os._exit()

subject = input("Enter subject: ")
body = input("Enter base body: ")
recipients = 'gamadavid36@gmail.com'
Emailer(body, subject, recipients)
