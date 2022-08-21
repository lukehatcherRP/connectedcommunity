# Simple script to read a list of scriptures in the format 'Genesis 1:1-12' 
# from a spreadsheet and output them in hyperlink form for Biblegateway.
# The structure is based on how the Biblegateway links are formatted which just repalces
# the emy ' ' space with a '+' and then tags teh versino at the end =ESV.
# Then at the bottom is is just formatted into 2 seperate out columns with full hpyerlink and short link

import pandas as pd
import xlwt 
from xlwt import *
import xlsxwriter

# Read the Workbook
workbook = pd.read_excel('Leviticus Series_SB.xlsx')
workbook.head()

# Create a new workbook and add a worksheet
workbook2 = xlsxwriter.Workbook('ConnectedCommunity.xls')
worksheet = workbook2.add_worksheet('Passages')


#https://www.biblegateway.com/passage/?search=Leviticus+14:1-20&version=ESV
item = True
i=0

while item == True:
    #print(workbook.iloc[i].iloc[1])
    text = workbook.iloc[i].iloc[1]
    #print(text)
    conv = text.replace(' ', '+')
    passage = "https://www.biblegateway.com/passage/?search=" + conv + "&version=ESV"
    #print(passage)
    
    worksheet.write_url(i, 1, passage, string=text)
    worksheet.write_url(i, 0, passage)
    i+=1
    try:
        workbook.iloc[i].item
    except:
        item = False
workbook2.close()

#wb.save('ConnectedCommunity.xls')
    