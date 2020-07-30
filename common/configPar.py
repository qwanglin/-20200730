# -*- coding: utf-8 -*- 
"""
# @Time : 2019/3/28
# @Author : xucaimin
"""

import configparser
import os
import sys
from common.read_conf import *


def getConfig(section, key):
    config = myconf()
    path =os.path.dirname(os.getcwd()) + '/config.conf'
    config.read(path, encoding="utf-8-sig")  # 此处是utf-8-sig，而不是utf-8
    return config.get(section, key)




if __name__=="__main__":
    print(getConfig("DEFAULT","url"))