# -*- coding: utf-8 -*- 

"""
# @Time : 2019/3/28 
# @Author : xucaimin
"""
from common.configPar import getConfig
from selenium import webdriver
import time
from common.WriteToTxt import logWriteToTxt
from common.zentaoInterface import _getiterate
from selenium.webdriver.common.action_chains import ActionChains
import os
import datetime
from common.configPar import *
from common.getfileName import *


url=getConfig("DEFAULT","url")

#打开网页,并且指定下载路径
def open_browser_login2():

    filepath = os.path.dirname(os.getcwd()) + "/file/" + datetime.datetime.now().strftime('%Y%m%d %H%M%S')
    print("创建的文件地址:",filepath)
    try:
        os.makedirs(filepath)
    except:
        pass
    #文件下载设置路径
    options = webdriver.ChromeOptions()
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': filepath}
    options.add_experimental_option('prefs', prefs)

    # 打开地址链接
    driver=webdriver.Chrome(chrome_options=options)
    # driver=webdriver.Chrome()
    driver.get(url+"/zentao/user-login.html")
    logWriteToTxt("进入到禅道登录页面" )
    driver.maximize_window() # 窗口最大化
    time.sleep(1)
    account = getConfig("DEFAULT", "account")
    password = getConfig("DEFAULT", "password")

    driver.find_element_by_id("account").send_keys(account)
    logWriteToTxt("输入用户名：" + account)
    driver.find_element_by_name("password").send_keys(password)
    logWriteToTxt("输入密码：" + password)
    driver.find_element_by_id("submit").click()
    logWriteToTxt("点击登录按钮" )
    time.sleep(1)#秒数
    return driver

#打开网页,但是不指定下载路径
def open_browser_login():
    filepath = os.path.dirname(os.getcwd()) + "/file/" + datetime.datetime.now().strftime('%Y%m%d %H%M%S')
    print("创建的文件地址:", filepath)
    try:
        os.makedirs(filepath)
    except:
        pass
    driver = webdriver.Chrome()
    driver.get(url+"/zentao/user-login.html")
    driver.maximize_window()  # 窗口最大化
    time.sleep(1)
    account = getConfig("DEFAULT", "account")
    password = getConfig("DEFAULT", "password")

    driver.find_element_by_id("account").send_keys(account)
    logWriteToTxt("输入用户名：" + account)
    driver.find_element_by_name("password").send_keys(password)
    logWriteToTxt("输入密码：" + password)
    driver.find_element_by_id("submit").click()
    logWriteToTxt("点击登录按钮")
    time.sleep(1)  # 秒数
    return driver



if __name__=="__main__":
    # open_browser_login2()
    print(os.getlogin())#知道电脑的用户名
    path="C:\\Users\\"+os.getlogin()+"\\Downloads"
    print(new_file(path))
    copy__file(new_file(path),"C:\\Users\\yanfa")




