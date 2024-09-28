import pandas as pd 
import re
##########################################
from gcsa.google_calendar import GoogleCalendar
from gcsa.event import Event
import datetime as dt
#########################################



calendar = GoogleCalendar("omarashrafmorshdy@gmail.com",credentials_path= ".\credentials\credentials.json" )
event = Event("Testing Event Headings",
              start=dt.datetime(2024,9,28,22,0),location="Online",minutes_before_popup_reminder= 30);
calendar.add_event(event)

for event in calendar:
    print(event)

df = pd.read_excel('sheet.xlsx')
# print(df)

# Added Columns from website to be deleted
droppedColumns = ['user_doc_id',
'modified_at_iso',
'status',
'review_url',
'uploaded_from',
'doc_meta_data',
'folder_name',
'folder_id',
'title',
'doc_id',
'created_at_iso',
'type']
# Deleting Added Columns from table
df = df.drop(columns=droppedColumns,axis=1)
# Removing added words in Columns
df.columns = df.columns.str.replace('Tables ','',regex=True)
# # Getting NaN Values
# nanValues = df.isna()

row  = 0
lst = [[]]

for x in range(df.shape[0]-2):
    for y in range(df.shape[1]):
        value = df.iloc[x,y]
        if not pd.isna(value):
            lst[row].append(df.columns[y])
            if type(value) == str and  re.match('[0-9]',str(value)):
                temp = value.split('-') 
                lst[row].append(temp[0])
                lst[row].append(temp[1])
            else:
                lst[row].append(value)
    lst.append([])
    row += 1

lst.pop()

for x in lst:
    if len(x)>4:
        print(x)      


        

