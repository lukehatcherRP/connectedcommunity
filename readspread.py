import xlwt 
from xlwt import *
import xlsxwriter

# Read the Workbook
workbook = xlrd.open_workbook('Leviticus Series_SB.xlsx')
sh = workbook.sheet_by_index(0)
#workbook.head()

# Create a new workbook and add a worksheet
workbook2 = xlsxwriter.Workbook('ConnectedCommunity.xls')
worksheet = workbook2.add_worksheet('Passages')


#https://www.biblegateway.com/passage/?search=Leviticus+14:1-20&version=ESV
#item = True
i=0

while i < 10:
    #print(workbook.iloc[i].iloc[1])
    text = sh(i,1)
    conv = text.replace(' ', '+')
    passage = "https://www.biblegateway.com/passage/?search=" + conv + "&version=ESV"
    print(passage)

    worksheet.write_url(i, 0, passage)
    i+=1

workbook2.close()

#wb.save('ConnectedCommunity.xls')
    