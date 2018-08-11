## 1. 批量下载DA文件  2. 合成一个文件

import xlwt
import pandas as pd
import os
import urllib3


# 下载文件函数
def DownLoad():
    http = urllib3.PoolManager()
    response = http.request('GET', url)
    with open(csv_file_name, 'wb') as f:
        f.write(response.data)

## 下载
# 设置文件名
year = '2018'
month = '06'

# https://marketplace.spp.org/file-api/download/da-lmp-by-bus?path=%2F2018%2F03%2FBy_Day%2FDA-LMP-B-201803010100.csv
url_basis = 'https://marketplace.spp.org/file-api/download/da-lmp-by-bus?path=%2F'+year+'%2F'+month+'%2FBy_Day%2FDA-LMP-B-'

for i in range(1, 32):
    k = 0
    if i < 10:
        day = '0' + str(i)
    else:
        day = str(i)
    url_part = year + month + day
    url = url_basis + url_part + '0100.csv'
    print(url)

    # 下载文件
    csv_file_name = 'DA-LMP-B-' + url_part + '0100.csv'
    if os.path.exists(csv_file_name):
        if not os.path.getsize(csv_file_name):
            DownLoad()
    else:
        DownLoad()


    #新建excel文件
    myexcel = xlwt.Workbook()
    sheet1 = myexcel.add_sheet(u'sheet1',cell_overwrite_ok=True)   # 创建sheet

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
                for a in range(0,12):
                    sheet1.write(k+a, 0, Time)
                    sheet1.write(k+a, 1, LMP)
                k = k + 12

        Savename = "F:\\Cirrus Wind 1\\研究DA和RT哪个好\\" + year + month + "_DA与RT对比\\" + url_part + ' - DayAhead.xls'
        myexcel.save(Savename)