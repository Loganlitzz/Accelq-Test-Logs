import win32com.client
import re
#other libraries to be used in this script
import os
from openpyxl import Workbook, load_workbook

from datetime import datetime, timedelta
outlook = win32com.client.Dispatch('outlook.application')
mapi = outlook.GetNamespace("MAPI")
inbox = mapi.GetDefaultFolder(6)
messages = inbox.Items
received_dt = datetime.now() - timedelta(days=1)
received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
messages = messages.Restrict("[SenderEmailAddress] = 'do-not-reply@accelq.com'")
#messages = messages.Restrict("[Subject] = 'Sample Report'")
message=messages.GetFirst()
body_content=message.body
result=re.search(r"Duration \d\d:\d\d:\d\d",body_content)
jobid=message.subject[50:62]

print(result.group())
print(received_dt)
print(jobid)

wb= load_workbook('Testrun_log.xlsx')
for sheet in wb:
    print(sheet)
ws=wb['Sheet1']
ws=wb.active
# ws['A1'].value='Duration'
print(ws['A1'].value)