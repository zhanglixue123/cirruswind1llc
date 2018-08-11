import xlwt
import pandas as pd
import os

k = 0

#新建excel文件
myexcel = xlwt.Workbook()
sheet1 = myexcel.add_sheet(u'sheet1',cell_overwrite_ok=True) #创建sheet

# 设置文件名
year = '2018'
month = '01'
day = '14'

for j in range(0,25):
    for i in range(0,13):

        if i < 2:
            minute = '0' + str(i*5)
        else:
            minute = str(i*5)
        if j < 10:
            hour = '0' + str(j)
        else:
            hour = str(j)

        if (i<12)&(j<24):
            csv_file_name = year + month + day + hour + minute + '.csv'
            # print(csv_file_name)

            #读取文件信息
            if os.path.getsize(csv_file_name):
                file_data = pd.read_csv(csv_file_name, usecols=['Interval', 'Pnode', 'LMP'], encoding='gbk')
                # 找cirruswind
                for find_i in range(0, file_data.shape[0]):
                    CIRRUSWIND = file_data.loc[find_i]
                    if 'CIRRUS' in CIRRUSWIND[1]:
                        Time = CIRRUSWIND[0]
                        Name = CIRRUSWIND[1]
                        LMP = CIRRUSWIND[2]
                        print('Time=',Time,'Name=',Name,'LMP=',LMP)
                        # 写入excel
                        # sheet1.write(k, 0, Time)
                        # sheet1.write(k+1, 0, Time)
                        # sheet1.write(k+2, 0, Time)
                        # sheet1.write(k+3, 0, Time)
                        # sheet1.write(k+4, 0, Time)
                        # sheet1.write(k + 5, 0, Time)
                        sheet1.write(k,1,LMP)
                        sheet1.write(k+1, 1, LMP)
                        sheet1.write(k+2, 1, LMP)
                        sheet1.write(k+3, 1, LMP)
                        sheet1.write(k+4, 1, LMP)
                        # sheet1.write(k + 5, 1, LMP)
                        k=k+5

                myexcel.save("20180114.xls")


