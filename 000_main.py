import easygui
import openpyxl
import sys
import os
import time
import ioMonitor
import cylinder
import hmiFile
#import gc

#input excel path
excelpath=easygui.enterbox(msg="请输入Excel文件完整路径")
start=time.perf_counter()
#open excel
try:
    wb1=openpyxl.load_workbook(filename=excelpath,read_only=True)
except:
    easygui.msgbox("Excel文件未找到，请检查路径名")
    sys.exit()

#open sheet
try:
    ws1=wb1["Work_Position"]
    ws2=wb1["Cylinder"]
    ws3=wb1["IO"]
except:
    easygui.msgbox("Excel表单未找到，请检查表单命名")
    sys.exit()
#D:\Project\Labview\test\python\Excel.xlsm
#create folder
rootpath=excelpath.rpartition('\\')
plcpath=rootpath[0]+r'\PLC'
hmipath=rootpath[0]+r'\HMI'
isExists=os.path.exists(plcpath)
if not isExists:
    os.mkdir(plcpath)
else:
    coverfolder=easygui.indexbox(msg='PLC文件夹已存在，是否覆盖？',choices=('是','否'))
    if coverfolder:
        sys.exit()
        
isExists=os.path.exists(hmipath)
if not isExists:
    os.mkdir(hmipath)
else:
    coverfolder=easygui.indexbox(msg='HMI文件夹已存在，是否覆盖？',choices=('是','否'))
    if coverfolder:
        sys.exit()

#create IO Sheet
if not ioMonitor.plcTag(ws3,plcpath):
    easygui.msgbox("IO表单填写不正确")
    sys.exit()

#create cylinder
cylinder.cylinder(ws2,plcpath,hmipath)

#create hmi files
hmiFile.hmiFile(ws1,ws2,ws3,hmipath)

#release memory
#del ws1,ws2,ws3,wb1
#gc.collect()

end=time.perf_counter()