# python-generate-plc-cylinder-program
使用python开发的软件工具，可以自动生成气缸的西门子PLC和HMI程序，大大提升工作效率  
B站视频地址：https://www.bilibili.com/video/BV1t44y1C7eD  
![image](https://github.com/siesen/python-generate-plc-program/blob/main/cover.PNG)
## 1.可以通过Eplan导出标签的方式导出IO信息，在Excel.xlsm中编辑IO以及气缸的中英文文本。（注：气缸IO需要黄色单元格，方便VBA识别是气缸的IO，工位名称后与气缸名称后加空格，方便VBA抽取数据）
## 2.在work_position和cylinder表单中点击按钮，运行VBA，抓取IO表单的数据，为Python做准备。
## 3.直接运行000_main.py，填入Excel.xlsm的绝对路径，点击确认。
## 4.在Excel.xlsm同目录下会生成PLC，HMI文件夹，里面的文件直接导入到西门子博途软件。
## 5.气缸的PLC，HMI程序就自动生成了，可以直接运行。
