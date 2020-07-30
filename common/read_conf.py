# -*- coding: utf-8 -*- 

"""
# @Time : 2019/3/28
# @Author : xucaimin
"""

import configparser
import os
import sys
from common.getfileName import getPath


# ConfigParser option 无论大小写，都会转换成小写,只能继承重写方法的方式
class myconf(configparser.ConfigParser):
    def __init__(self,defaults=None):
        configparser.ConfigParser.__init__(self,defaults=None)
    def optionxform(self, optionstr):
        return optionstr

def readconfig():
    path = getPath() + "/config.conf"

    print(1, getPath())
    path2 = os.path.abspath(os.path.dirname(__file__))
    print(2, path2)
    print(3, sys.path[0])
    print(4, os.path.split(sys.path[0])[0])
    print(5,os.path.dirname(os.getcwd()))
    path = os.path.dirname(os.getcwd()) + "/config.conf"
    print("6",path)

    print(1)
    conf = configparser.ConfigParser()
    conf.read(path, encoding="utf-8-sig")  # 此处是utf-8-sig，而不是utf-8
    print(conf.sections())
    for i in conf.sections():
        print(conf.options(i))
        for option in conf.options(i):
            print(option, conf.get(i, option))

    print(2)
    conf = myconf()
    conf.read(path, encoding="utf-8-sig")  # 此处是utf-8-sig，而不是utf-8
    print(conf.sections())
    for i in conf.sections():
        print(conf.options(i))
        for option in conf.options(i):
            print(option, conf.get(i, option))


if __name__=="__main__":
    readconfig()



