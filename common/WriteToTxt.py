# -*- coding: utf-8 -*- 
"""
# @Time : 2019/3/28
# @Author : xucaimin
"""
import os
import datetime
import sys

def logWriteToTxt(data):

    try:
        os.makedirs(os.path.dirname(os.getcwd()) + "/Data/")
    except:
        pass

    #向该项目下的/Data/Log的当前时间的log文档（如没文件不存在则创建）写入日志data
    path = os.path.dirname(os.getcwd()) + "/Data/Log" + datetime.datetime.now().strftime('%Y-%m-%d') + "log.txt"
    with open(path, "a+", encoding="utf-8") as f:
        if(data!=""):
            f.write(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + "   " + data + "\n")
        else:
            f.write("\n")

def getPath():
    path=sys.path[0]#如果PYTHONPATH 变量还不存在，可以创建它！路径会自动加入到sys.path中
    return path

if __name__ == "__main__":
    path = os.path.abspath(os.path.dirname(__file__))
    type = sys.getfilesystemencoding()

    print(path)
    print(os.path.dirname(__file__))
    print('------------------')
    print(os.path.split(sys.path[0])[0])




