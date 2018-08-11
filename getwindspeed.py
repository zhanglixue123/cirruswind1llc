# 把下载的txt风速转为csv
# csv数据读取到xls里，合并，算5min平均

import pandas as pd
import xlwt
import os
import csv

YEAR = '2018'
MONTH = '06'

os.chdir("F:\\Cirrus Wind 1\\研究DA和RT哪个好\\" + YEAR + MONTH + "_DA与RT对比\\windspeed")

# TXT转为csv
for file in os.listdir():
    thisFile = file
    base = os.path.splitext(thisFile)[0]
    ext = os.path.splitext(thisFile)[1]
    # print(ext)
    if ext == '.txt':
        pd.read_csv(thisFile,sep='\t',engine='python').to_csv(base+'.csv')
        # os.remove(thisFile)

# 写入excel
#新建excel文件
myexcel = xlwt.Workbook()

a = 0
for i in range(0, 31):   # 日期
    sheetname = str(i+1)
    sheet1 = myexcel.add_sheet(sheetname, cell_overwrite_ok=True)  # 创建sheet
    for j in range(0, 17):  # 风机
        csv_file_name = YEAR + MONTH + "_" + str(i+1) + "_" + str(j+1) + "#.csv"
        if os.path.exists(csv_file_name):
            if os.path.getsize(csv_file_name):
                # 读取风速
                with open(csv_file_name, "rt", encoding="utf-8") as csvfile:
                    reader = csv.reader(csvfile)
                    column = [row[3] for row in reader]
                with open(csv_file_name, "rt", encoding="utf-8") as csvfile:
                    reader = csv.reader(csvfile)
                    Time = [row[2] for row in reader]
                print(column)
                print(Time)
                # 写入excel
                for a in range(0, len(column)):
                    sheet1.write(a, j+1, float(column[a]))     # 写风速
                    sheet1.write(a, 0, Time[a])     # 写日期时间
                    WINDSPEED_AVE = 'AVERAGE(P' + str(a+1) + ':R' + str(a+1) + ',B' + str(a+1) + ':N' +str(a+1) + ')'
                    sheet1.write(a, 18, xlwt.Formula(WINDSPEED_AVE))    # 计算所有风机平均风速
                for b in range(0, 288):
                    WINDSPEED_AVE_5min = 'AVERAGE(S' + str(1+b*5) + ':S' + str(1+b*5+4) + ')'
                    sheet1.write(b, 19, xlwt.Formula(WINDSPEED_AVE_5min))  # 计算5分钟平均风速


Savename = YEAR + MONTH + ' - 风速.xls'
myexcel.save(Savename)

os.chdir("C:\\Users\\Dell\\PycharmProjects\\sherry")

