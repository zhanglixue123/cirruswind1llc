import pandas as pd
import urllib3
#
# 设置文件名
year = '2018'
month = '01'
day = '14'
url_basis = 'https://marketplace.spp.org/file-api/download/rtbm-lmp-by-bus?path=%2F'+year+'%2F'+month+'%2FBy_Interval%2F' + day +'%2FRTBM-LMP-B-'
for j in range(0,25):
    for i in range(0,13):
        # print("i=", i, "j=", j)

        if i < 2:
            minute = '0' + str(i*5)
        else:
            minute = str(i*5)
        if j < 10:
            hour = '0' + str(j)
        else:
            hour = str(j)

        if (i<12)&(j<24):
            url_part = year + month + day + hour + minute
            url = url_basis + url_part + '.csv'
            print(url)

            # 下载文件
            csv_file_name = year + month + day + url_part + '.csv'
            http = urllib3.PoolManager()
            response = http.request('GET', url)
            with open(url_part + '.csv', 'wb') as f:
                f.write(response.data)


