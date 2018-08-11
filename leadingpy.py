import sys

PYname = input('请输入程序名称：')

while(1):
    with open(PYname, 'rb') as f:
        code = f.read().decode()
        try:
            if sys.version_info.major == 2:
                exec(code)
            else:
                exec(code)
        except:
            print('failed')
            # os.system("start C:\Biexiangta.mp3")
            # os.system("start C:\程序意外终止.txt")
        # else:
        #     print('success')
        #     os.system("start C:\Biexiangta.mp3")
        #     os.system("start C:\程序已运行完毕.txt")