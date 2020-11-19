import win32com.client
import pandas
import datetime
import re
from site_info import SiteInfo

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

site = SiteInfo()
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
msg = messages.GetLast()
msg_date = re.search(r"\d\d\d\d-\d\d-\d\d", str(msg.SentOn))
today = str(datetime.date.today())
subject = []
body = []

while msg_date.group(0) == today:
    subject.append(msg.subject+site.get_domain())
    body.append(re.sub(r'\r\n\r\n', r'\r\n', msg.body))
    msg = messages.GetPrevious()
    msg_date = re.search(r"\d\d\d\d-\d\d-\d\d", str(msg.SentOn))

    if msg_date.group(0) != today:
        break

df = pandas.DataFrame()
df['Subject'] = subject
df['Body'] = body
outlook = win32com.client.Dispatch('outlook.application')

# Send out all emails received
for i in range(len(df)):
    mail = outlook.CreateItem(0)
    mail.To = df['Subject'].iloc[i]
    mail.Subject = "You've got mail! Anonymous Recognition!"
    mail.Body = df['Body'].iloc[i]
    mail.Display(True)
    # mail.Send()
