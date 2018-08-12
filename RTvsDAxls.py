# 生成RT vs DA的xls模板

import xlwt
from xlwt import *
import xlrd
import os

YEAR = '2018'
MONTH = '06'
Real_Time_Support = "F:\\Cirrus Wind 1\\xcelenergy\\2018-06 Cirrus Support.xlsx"
Support_sheet = '2018-06 Cirrus Support'
Day_Ahead_combine = '201806 - 合并.xlsx'
combine_sheet = 'Sheet1'

os.chdir("F:\\Cirrus Wind 1\\研究DA和RT哪个好\\" + YEAR + MONTH + "_DA与RT对比")


#新建excel文件
myexcel = xlwt.Workbook()
sheetname = YEAR + MONTH
sheet = myexcel.add_sheet(sheetname,cell_overwrite_ok=True)   # 创建sheet

#### 模板

style = XFStyle()
fnt = Font()
fnt.bold = True
style.font = fnt
# sheet.write_merge(first row, last row, first col, last col, 单元格value, 格式style)
sheet.write_merge(0, 0, 0, 2, 'REAL TIME', style)
sheet.write_merge(0, 0, 4, 6, 'DAY AHEAD', style)
sheet.write(0, 7, '风速', style)
sheet.write(1, 0, 'DATE')
sheet.write(1, 1, 'HOUR')
sheet.write(1, 2, 'LMP($/MWh)')
sheet.write(1, 4, 'TIME')
sheet.write(1, 5, 'LMP($/MWh)')
sheet.write(1, 7, 'm/s')
sheet.write(1, 8, '>9')
sheet.write(1, 9, '>8')
sheet.write(1, 10, '>7')
sheet.write(1, 11, 'Price D')
sheet.write(1, 13, 'Wind>7')
sheet.write(2, 13, 'Average')
sheet.write(3, 13, 'Sum')

# 公式 excel formulas
for i in range(3, 8931):
    MoreThan9 = 'IF(H' + str(i) + '>9,"√","")'      # >9
    sheet.write(i-1, 8, xlwt.Formula(MoreThan9))
    # print(MoreThan9)
    MoreThan8 = 'IF(H' + str(i) + '>8,"√","")'      # >8
    sheet.write(i-1, 9, xlwt.Formula(MoreThan8))
    # print(MoreThan8)
    MoreThan7 = 'IF(H' + str(i) + '>7,"√","")'      # >7
    sheet.write(i-1, 10, xlwt.Formula(MoreThan7))
    # print(MoreThan7)
    PriceD = 'IF(H' + str(i) + '>7,C' + str(i) + '-F' + str(i) + ',"")'  # Price D
    sheet.write(i - 1, 11, xlwt.Formula(PriceD))
sheet.write(2, 14, xlwt.Formula('AVERAGE(L3:L8930)'))
sheet.write(3, 14, xlwt.Formula('SUM(L3:L8930)'))


### 导入数据 
# real time
workbook = xlrd.open_workbook(Real_Time_Support)
booksheet = workbook.sheet_by_name(Support_sheet)
table = workbook.sheets()[0]
nrows = table.nrows    #行数
# print(nrows)
for i in range(1, nrows):
    cell_DATE = booksheet.cell_value(i, 1)
    cell_HOUR = booksheet.cell_value(i, 2)
    cell_LMP = booksheet.cell_value(i, 9) * 1000
    # print(str(cell_DATE)+"  "+str(cell_HOUR))
    sheet.write(i+1, 0, cell_DATE)
    sheet.write(i+1, 1, cell_HOUR)
    sheet.write(i+1, 2, cell_LMP)
# day ahead
workbook = xlrd.open_workbook(Day_Ahead_combine)
booksheet = workbook.sheet_by_name(combine_sheet)
table = workbook.sheets()[0]
nrows = table.nrows    #行数
for i in range(0, nrows):
    cell_TIME = booksheet.cell_value(i, 0)
    cell_DA = booksheet.cell_value(i, 1)
    sheet.write(i+2, 4, cell_TIME)
    sheet.write(i+2, 5, cell_DA)

# 粘贴风速 paste wind speeds
Windspeedxls = "F:\\Cirrus Wind 1\\研究DA和RT哪个好\\" + YEAR + MONTH + "_DA与RT对比\\windspeed\\" + YEAR + MONTH + " - 风速.xls"
workbook = xlrd.open_workbook(Windspeedxls)
for i in range(0, len(workbook.sheets())):
    booksheet = workbook.sheet_by_index(i)
    table = workbook.sheets()[i]
    for j in range(0, 288):
        WINDSPEED_AVE_5min = booksheet.cell_value(j, 19)
        if WINDSPEED_AVE_5min == 7:
            sheet.write(j + 288 * i + 2, 7, '')
        else:
            sheet.write(j + 288 * i + 2, 7, WINDSPEED_AVE_5min)




# 保存文件
Savename = YEAR + MONTH + ' - RTvsDA.xls'
myexcel.save(Savename)
