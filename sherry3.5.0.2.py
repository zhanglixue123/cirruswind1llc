# 2018.3.12更新：下载改为5分钟一次，下载前先检测文件是否已存在。新增timeout功能防止下载假死。
# 2018.3.4更新：读取时间部分大改，改为用time.time()命令更改时间，缩短代码行数。功能不变。
# 2018.1.10更新，新添“开始运行”
# 2018.1.5更新，新添自动操作鼠标调功率
import urllib3
import socket
import pandas as pd
import time
import os
import itchat
import win32api
import win32con
from ctypes import *

# 定义剪切板函数
def addToClipBoard(text):
    command = 'echo ' + text.strip() + '| clip'
    os.system(command)

# 定义自动修改功率函数
def AutoChangeMW():
    # 打开teamviewer
    windll.user32.SetCursorPos(459, 744)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(0.05)
    # 关闭对话框
    windll.user32.SetCursorPos(855, 417)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    windll.user32.SetCursorPos(855, 417)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(0.5)
    windll.user32.SetCursorPos(785, 402)  # 新电脑
    # windll.user32.SetCursorPos(780, 392) #老电脑
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    windll.user32.SetCursorPos(785, 402)  # 新电脑
    # windll.user32.SetCursorPos(780, 392) #老电脑
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(0.8)
    # 打开功率设置
    windll.user32.SetCursorPos(370, 130)  # 新电脑
    # windll.user32.SetCursorPos(313, 145)  # 老电脑
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(0.5)
    # 选中
    windll.user32.SetCursorPos(655, 366)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    # 删掉原来的功率
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    win32api.keybd_event(8, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(8, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.25)
    ## 记录调节过程
    try:
        RecordFile = open(RecordFileName, 'a')
    except IOError:
        print('*** file open error:')
    else:
        RecordFile.write('\n' + str(time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time()))) + '\n')
        if LMP >= 0:
            for Q in range(0, 3):
                addToClipBoard('61200')  # 把剪切板设置为61200
            RecordFile.write('新的电价： ' + str(LMP) + '，现在请调整出力为61200' + '\n')
        else:
            for Q in range(0, 3):
                addToClipBoard('2000')  # 把剪切板设置为2000
            RecordFile.write('新的电价： ' + str(LMP) + '，现在请调整出力为2000' + '\n')
        RecordFile.write('下载文件：' + csv_file_name + '\n')
        RecordFile.close()
    # 修改功率
    win32api.keybd_event(17, 0, 0, 0)  # ctrl键位码是17
    win32api.keybd_event(86, 0, 0, 0)  # v键位码是86
    win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放按键
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.5)
    # 点击确定
    windll.user32.SetCursorPos(726, 364)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(0.2)
    # 最小化teamviewer窗口
    windll.user32.SetCursorPos(459, 744)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

# 定义下载文件函数
def DownLoad():
    try:
        headers = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36'}
        timeout = 90
        socket.setdefaulttimeout(timeout)
        http = urllib3.PoolManager()
        response = http.request('GET', url, headers=headers)
        with open(url_part + '.csv', 'wb') as f:
            f.write(response.data)
        print(csv_file_name)
    except:
        print("下载文件失败！！！")
        os.system("start C:\Biexiangta.mp3")
        os.system("start C:\下载文件失败.txt")
        pass

# 定义发微信函数
def SendWeChat(message):
    try:
        itchat.send(message, toUserName=username)
    except:
        print("发送微信失败！！！")
        os.system("start C:\Biexiangta.mp3")
        os.system("start C:\发送微信失败.txt")
        pass

# # 设置微信好友
# itchat.login()
# users = itchat.search_friends(name = '宇杉😘Sherry')  # 宇杉😘Sherry
# username = users[0]['UserName']

# 开始运行
LMP = -1111111111
record = 3
delay = 0
Delay = 0
DIFF = 3
R_csv_file_name = 'start'
recover = 0
BadPing = 0
FileBuild = 0
while(1):
    # 新建记录文件
    RecordFileName = time.strftime('%Y%m%d', time.localtime(time.time())) + '_程序运行记录' + '.txt'
    if not os.path.exists(RecordFileName):
        if FileBuild == 0:
            RecordFile = open(RecordFileName, 'w')
            RecordFile.close()
            try:
                RecordFile = open(RecordFileName, 'a')  # 这里的a意思是追加，这样在加了之后就不会覆盖掉源文件中的内容，如果是w则会覆盖。
            except IOError:
                print('*** file open error:')
            else:
                RecordFile.write('\n' + '\n' + '***  ' + str(time.strftime('%Y-%m-%d', time.localtime(time.time()))) + '  电价监测程序运行记录' + '  ***' + '\n' + '\n')
                RecordFile.close()  # 特别注意文件操作完毕后要close
            FileBuild = 1
    # Part I: make the file name ####
    local_time = time.localtime(time.time())
    YEAR = time.strftime('%Y', local_time)
    MONTH = time.strftime('%m', local_time)
    DAY = time.strftime('%d', local_time)
    url_basis = 'https://marketplace.spp.org/file-api/download/rtbm-lmp-by-bus?path=%2F'+YEAR+'%2F'+MONTH+'%2FBy_Interval%2F' + DAY +'%2FRTBM-LMP-B-'
    # 改时间
    second = float(time.strftime('%S', local_time))
    minute = (int(time.strftime('%M', local_time))) % 5
    if ((second > 13) & (second <= 17)) | (((second > 41) & (second <= 44))&(minute < 3)): # 一分钟只下载两次
        url_part = time.strftime('%Y%m%d%H%M', time.localtime(time.time() - (time.time() % 300) + 300)) # min+5
        url = url_basis + url_part + '.csv'
        csv_file_name = url_part + '.csv'
    # Part II: start the python engnine
        exit_code = os.system('ping www.google.com')
        if exit_code:
            if BadPing == 0:
                os.system("start C:\Biexiangta.mp3")
                os.system("start C:\网络连接已断开.txt")
                try:
                    RecordFile = open(RecordFileName, 'a')  # 这里的a意思是追加，这样在加了之后就不会覆盖掉源文件中的内容，如果是w则会覆盖。
                except IOError:
                    print('*** file open error:')
                else:
                    RecordFile.write('\n' + str(time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time()))) + '\n')
                    RecordFile.write('网络连接已断开，等待重连……' + '\n')
                    RecordFile.close()
                BadPing = BadPing + 1
            time.sleep(6)
        else:
            BadPing = 0
            if os.path.exists(csv_file_name):
                if not os.path.getsize(csv_file_name):
                    DownLoad()
            else:
                DownLoad()

    # Part IV: Post-process the downloaded file
            # 读取csv
            if os.path.getsize(csv_file_name):
                file_data = pd.read_csv(csv_file_name, usecols=['Interval', 'Pnode', 'LMP'], encoding='gbk')
                # 找cirruswind
                for find_i in range(0, file_data.shape[0]):
                    CIRRUSWIND = file_data.loc[find_i]
                    if 'CIRRUS' in CIRRUSWIND[1]:
                        Time = CIRRUSWIND[0]
                        Name = CIRRUSWIND[1]
                        LMP = CIRRUSWIND[2]
                # 判断LMP有无正负变化
                if LMP >= 0:
                    judgement = 1
                else:
                    judgement = 0
                diff = record - judgement
                print('record=', record, 'judgement=', judgement, 'diff =', diff)
                if record == 3:
                    try:
                        RecordFile = open(RecordFileName, 'a')
                    except IOError:
                        print('*** file open error:')
                    else:
                        RecordFile.write('\n' + str(time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time()))) + '\n')
                        if LMP >= 0:
                            message = '现在开始运行，目前电价： ' + str(LMP) + '。请注意保持出力为61200'
                            SendWeChat(message)
                            RecordFile.write('现在开始运行，目前电价：' + str(LMP) + '。请注意保持出力为61200' + '\n')
                        else:
                            message = '现在开始运行，目前电价： ' + str(LMP) + '。请注意保持出力为2000'
                            SendWeChat(message)
                            RecordFile.write('现在开始运行，目前电价：' + str(LMP) + '。请注意保持出力为2000' + '\n')
                        RecordFile.write('下载文件：' + csv_file_name + '\n')
                        RecordFile.close()
                    AutoChangeMW()  # 防止程序重启后无法判断
                if (judgement != record) | ((diff == 0) & (DIFF != 3) & (delay != 0) & (DIFF != 2)&(R_csv_file_name != csv_file_name)):
                    if record != 3:
                        recover = 0
                        try:
                            RecordFile = open(RecordFileName, 'a')
                        except IOError:
                            print('*** file open error:')
                        else:
                            RecordFile.write('\n' + str(time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time()))) + '\n')
                            if judgement != record:
                                message = '有正负变动，新的电价： ' + str(LMP)
                                SendWeChat(message)
                                RecordFile.write('有正负变动，新的电价： ' + str(LMP) + '\n')
                            else:
                                message = '持续观察，新的电价： ' + str(LMP)
                                SendWeChat(message)
                                RecordFile.write('持续观察，新的电价： ' + str(LMP) + '\n')
                            RecordFile.write('下载文件：' + csv_file_name + '\n')
                            RecordFile.close()
                        # 还原
                        if (DIFF != diff) & (diff != 0) & (DIFF != 2) & (DIFF != 3) & (Delay < 4) & (delay > 0):
                            delay = 0
                            print('还原，delay=', delay)
                            recover = 1
                            message = '请忽略上次的变动。本次电价： ' + str(LMP)
                            SendWeChat(message)
                            try:
                                RecordFile = open(RecordFileName, 'a')
                            except IOError:
                                print('*** file open error:')
                            else:
                                RecordFile.write('\n' + str(time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time()))) + '\n')
                                RecordFile.write('请忽略上次的变动。本次电价： ' + str(LMP) + '\n')
                                RecordFile.write('下载文件：' + csv_file_name + '\n')
                                RecordFile.close()
                        print('record=', record, 'judgement=', judgement, 'diff =', diff)
                        # delay intervals
                        if recover == 0:
                            if abs(LMP) <= 2.5:
                                delay = delay + 0.5
                                print('delay=', delay)
                            elif (abs(LMP) > 2.5) & (abs(LMP) <= 5):
                                delay = delay + 1
                                print('delay=', delay)
                            elif abs(LMP) > 10:
                                delay = delay + 4
                                print('delay=', delay)
                            else:
                                delay = delay + 2
                                print('delay=', delay)
                        # 警报
                        Delay = delay
                        print('Delay=', Delay)
                        if Delay >= 4:
                            print('delay=', delay)
                            print('delay>=4,播放音乐')
                            os.system("start C:\shimian.mp3")  # 报警
                            delay = 0
                            AutoChangeMW()   # 自动打开teamviewer改出力
                            if LMP >= 0:
                                message = '新的电价： ' + str(LMP) + '，现在请调整出力为61200'
                            else:
                                message = '新的电价： ' + str(LMP) + '，现在请调整出力为2000'
                            SendWeChat(message)  # 报警
                elif R_csv_file_name != csv_file_name:
                    try:
                        RecordFile = open(RecordFileName, 'a')  # 这里的a意思是追加，这样在加了之后就不会覆盖掉源文件中的内容，如果是w则会覆盖。
                    except IOError:
                        print('*** file open error:')
                    else:
                        RecordFile.write('\n' + str(time.strftime('%Y-%m-%d  %H:%M:%S', time.localtime(time.time()))) + '\n')
                        RecordFile.write('无操作。现在电价为：' + str(LMP) + '\n')
                        RecordFile.write('下载文件：' + csv_file_name + '\n')
                        RecordFile.close()
                record = judgement
                DIFF = diff
                R_csv_file_name = csv_file_name
                print('LMP=', LMP)
                print('record=', record, 'judgement=', judgement)
                print('DIFF=', DIFF)
                print('Delay=', Delay)
                print('recover=', recover)
                time.sleep(2)
    print(time.strftime("%Y-%m-%d   %H:%M:%S", time.localtime(time.time())), '  ', 'LMP=', LMP, '   ', 'Delay=', Delay, '   ', 'recover=', recover)

