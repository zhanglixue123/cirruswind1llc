import win32clipboard as w
import win32con
import win32api
import xlwt
import os
import time
from ctypes import *


YEAR = '2018'
MONTH = '06'



# 点击 click
def ClickLeft():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(0.5)

# 点击风机号 click machine number 
def ClickFanNum():
    windll.user32.SetCursorPos(240, 155)
    ClickLeft()

# 点击滚动条 向上 click scroll bar upward
def ReturnUp():
    for k in range(0, 9):  # range(0, 9)
        windll.user32.SetCursorPos(251, 166)
        ClickLeft()

# 点击获取 get
def ClickGet():
    windll.user32.SetCursorPos(290, 155)
    ClickLeft()

# 全选+复制 alt+A, ctl+C
def CtrlAandCtrlC():
    windll.user32.SetCursorPos(785, 317)
    ClickLeft()
    win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
    win32api.keybd_event(65, 0, 0, 0)  # a键位码是65
    win32api.keybd_event(65, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
    win32api.keybd_event(67, 0, 0, 0)  # c键位码是67
    win32api.keybd_event(67, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.7)

# 保存文件 save 
def SaveTXT():
    with open(filename, 'w') as fileobject:  # 使用‘w’来提醒python用写入的方式打开
        fileobject.write(sss.decode('utf-8'))
    fileobject.close()

# 打开teamviewer
windll.user32.SetCursorPos(459, 744)
ClickLeft()

# 复制数据 locations of dates
# Date = [165, 175, 185, 195, 205, 215, 225, 235, 245, 249,
#         259, 269, 279, 289, 299, 309, 319, 329, 339, 349,
#         359, 369, 379, 389, 390, 400, 410, 420, 430, 440,
#         450]
Date = [329, 339, 349, 359, 369, 379, 389, 390, 400, 410,
        420, 430, 440, 450, 460, 465, 475, 485, 495, 505,
        515, 525, 535, 545, 555, 565, 575, 585, 595, 599,
        609]

# 新建windspeed文件夹
cur_dir = "D:/Cirrus Wind 1/研究DA和RT哪个好/" + YEAR + MONTH + "_DA与RT对比/"
folder_name = 'windspeed'
# if os.path.isdir(cur_dir):
if not os.path.exists(cur_dir + folder_name):
    os.mkdir(os.path.join(cur_dir, folder_name))
os.chdir("D:\\Cirrus Wind 1\\研究DA和RT哪个好\\" + YEAR + MONTH + "_DA与RT对比\\windspeed")

for i in range(0, 30):
    # 点击查询日期 click the date
    windll.user32.SetCursorPos(140, Date[i])
    ClickLeft()
    # 查看是否存在文件
    count = 1
    for m in range(1, 18):
        filename = YEAR + MONTH + "_" + str(i + 1) + "_" + str(m) + "#.txt"
        if os.path.exists(filename):
            count += 1
    if count < 18:
        # 风机 1~8 号 machine #1~#8
        for j in range(0, 8):   # range(0, 8)
            # 点击风机号 click the machine # 
            ClickFanNum()
            if j == 0 :
                # 点击滚动条 向上 click scroll bar upward
                ReturnUp()
            # 选择风机 choose machine
            windll.user32.SetCursorPos(240, 170 + j*11)
            ClickLeft()
            # 点击获取 get
            ClickGet()
            # 全选+复制 atl+A copy
            CtrlAandCtrlC()
            # 读取剪贴板内容 clipboard
            w.OpenClipboard()
            sss = w.GetClipboardData(win32con.CF_TEXT)
            w.CloseClipboard()
            # 新建txt文件
            filename = YEAR + MONTH + "_" + str(i+1) + "_" + str(j+1) + "#.txt"
            # 保存文件
            SaveTXT()
        # 风机 9~17 号 fan #1~#8
        for j in range(0, 9): # range(0, 9)
            # 点击风机号
            ClickFanNum()
            # 点击滚动条 向下
            windll.user32.SetCursorPos(251, 253)
            ClickLeft()
            # 点击风机 9 号
            windll.user32.SetCursorPos(240, 253)
            ClickLeft()
            # 点击获取
            ClickGet()
            # 全选+复制
            CtrlAandCtrlC()
            # 读取剪贴板内容
            w.OpenClipboard()
            time.sleep(0.25)
            sss = w.GetClipboardData(win32con.CF_TEXT)
            time.sleep(0.5)
            w.CloseClipboard()
            time.sleep(0.05)
            # 新建txt文件
            filename = YEAR + MONTH + "_" + str(i+1) + "_" + str(j+9) + "#.txt"
            # 保存文件
            SaveTXT()

os.chdir("C:\\Users\\zhang\\PycharmProjects\\sherry")
