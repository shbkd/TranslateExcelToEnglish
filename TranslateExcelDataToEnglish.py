## Get your google developer key
#  It will look something like below
#  developerKey = 'AIzaSyBC576lJHaERzcoFM1kRuuWqDlKp00HON6'
yourKey = 'AIzaSyBC576lJHaERzcoFM1kRuuWqDlKp00HON6'

#  Firstly, import relevant data analytics libraries
import pandas as pd
import numpy as np
import unicodedata
from pandas import ExcelWriter
from googleapiclient.discovery import build

#  bind yourself to api
service = build('translate', 'v2',
            developerKey=yourKey)

#  Now, Load the excel data into a panda data_frame
df = pd.read_excel('DiDi_Raw_SWO_Data_translation.xlsm', sheetname='SWO DXR')

#  Query data_frame to see structure of your data
print df.info()

#  Select Interested columns from this data-set
df = df[['SWO', 'CustomerComplaintSubject', 'CustomerComplaint', 'InternalRepairText', 'ExternalRepairText']]

#  Translate
for i in range(1785):
    print i
    cell_text = df['CustomerComplaintSubject'][i]
    #check if not nan
    if not pd.isnull(cell_text):
        # if encoding doesn't result in error
        try:
            translate = unicodedata.normalize('NFKD', cell_text).encode('ascii','ignore')
        except Exception:
            translate = cell_text
            continue
        translated = service.translations().list(
                     target='en',
                     q=[translate]
                     ).execute()

        df['CustomerComplaintSubject'][i] = translated['translations'][0]['translatedText']

#  Write output data back to excel
writer = ExcelWriter('output.xlsx')
df.to_excel(writer,'Sheet1')
writer.save()