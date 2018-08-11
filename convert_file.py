import os
import pandas

os.chdir("C:\\Users\\Dell\\PycharmProjects\\sherry\\windspeed")
for file in os.listdir():
    thisFile = file
    base = os.path.splitext(thisFile)[0]
    ext = os.path.splitext(thisFile)[1]
    # print(ext)
    if ext == '.txt':
        pandas.read_csv(thisFile,sep='\t',engine='python').to_csv(base+'.csv')
        os.remove(thisFile)