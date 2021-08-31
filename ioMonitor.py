# -*- coding: utf-8 -*-
import openpyxl
#create IO monitor
def plcTag(ws3,plcpath):
       
    #create PLC tag---------------------------------------------------------
    wb2=openpyxl.Workbook()
    ws4=wb2.active
    ws4.title='PLC Tags'
    ws4.append(['Name','Path','Data Type','Logical Address','Comment','Hmi Visible','Hmi Accessible'])
    for row in ws3.values:
        if row[0]:
            if row[0].startswith('I'):
                ws4.append([row[1],'IO表','Bool','%'+row[0],'','True','True'])
    for row in ws3.values:
        if row[4]:
            if row[4].startswith('Q'):
                ws4.append([row[5],'IO表','Bool','%'+row[4],'','True','True'])
    wb2.save(plcpath+"\\PLCTags.xlsx")
    wb2.close()
    return 1