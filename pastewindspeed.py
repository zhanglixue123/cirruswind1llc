# 查查哪个sheet的行数不够，补充

import xlrd
import xlwt
import os

YEAR = '2018'
MONTH = '06'

# 查查哪个sheet的行数不够
Windspeedxls = "F:\\Cirrus Wind 1\\研究DA和RT哪个好\\" + YEAR + MONTH + "_DA与RT对比\\windspeed\\" + YEAR + MONTH + " - 风速.xls"
workbook = xlrd.open_workbook(Windspeedxls)
for i in range(0, len(workbook.sheets())):
    booksheet = workbook.sheet_by_index(i)
    table = workbook.sheets()[i]
    nrows = table.nrows  # 行数
    if nrows < 1436:
        print("数据不全的sheet：" + str(i+1))
    cell = booksheet.cell_value(287, 19)
    if cell == 7:
        print("S288是空值的sheet：" + str(i+1))

# # 粘贴风速
#     for j in range(0, 288):
#         WINDSPEED_AVE_5min = booksheet.cell_value(j, 19)
#


