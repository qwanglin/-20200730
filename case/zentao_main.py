# -*- coding: utf-8 -*- 

"""
# @Time : 2019/3/29 
# @Author : xucaimin
"""
import sys
sys.path.append("../")
from common.configPar import *
from selenium import webdriver
import time
from common.WriteToTxt import *
from common.zentaoInterface import _login
from selenium.webdriver.common.action_chains import ActionChains
import os
import datetime
from common.zentao_login import *
from case.zentao_getTaskBugFiles import *


class Logger(object):
  def __init__(self, filename="Default.log"):
    self.terminal = sys.stdout
    self.log = open(filename, "a")
  def write(self, message):
    self.terminal.write(message)
    self.log.write(message)
  def flush(self):
    pass


if __name__=="__main__":

    fpath=os.path.split(sys.path[0])[0]
    sys.stdout = Logger(fpath+'/logger.txt')
    today = time.strftime("%Y%m%d %H%M%S")  # 今天
    print("执行时间为："+today+ " "+fpath)
    isHalfYear = getConfig("DEFAULT", "isHalfYear")
    #1.登录打开网页
    driver=open_browser_login()
    #2.得到迭代的文件--下载到本地（指定目录)
    session=getTaskFiles(driver,isHalfYear)
    # print("session:"+session)
    #3.修改csv文件
    change_task_csv(session)
    #4.合并csv文件并转化为excel
    taskxlsxpath = merge_csv("tasks",session,0,False)
    # 第1个sheet：前一周已完成的任务和所有未完成的任务、未开始和进行中
    getWeekTasksExcel()#得到上一周的任务数据（已完成的任务和所有未完成的任务、未开始和进行中--截止时间小于）
    #第2个sheet：前一月已完成的任务和所有未完成的任务
    getMonthTasksExcel(0)  # 得到上一月的任务考核数据



    # elif (today_day == "26"):


    if (int(isHalfYear) != 0):
        getClosedTaskFiles(driver, session, isHalfYear)
        # 再一次修改
        change_task_csv(session)
        # 再一次合并修改
        merge_csv("tasks", session,isHalfYear,False)

    # 1.登录打开网页
    driver = open_browser_login()
    # 2.得到bug文件--下载到本地（指定目录
    getBugFiles(driver)
    # 3.修改csv文件
    change_bug_csv(session)
    # 4.合并csv文件，并转化为excel
    merge_csv("bugs",1,0,False)
    # 第3个sheet：上月已解决的BUG和截止至导出数据时需要解决而未解决的所有的BUG
    getMonthBugsSolveExcel(0)
    # 第4个sheet：上月已回归的BUG和截止至导出数据时需要回归而未回归的所有的BUG。
    getMonthBugsReturnExcel(session,0)
    # 第5个sheet：上月已创建的bug
    getMonthBugsCreateExcel(session,0)
    #上月已关闭的bug
    getMonthBugsCloseExcel(session,0)

    # 获得版本测试单，用例测试单
    getResponVersionTestExcel(session)  #????
    getExcuteVersionTestExcel(session)
    # 获得上个月的版本测试单--得到两个excel
    getMonthTestExcel(0)
    # 最后最后合并tasks,bugs,版本测试单
    # 并且时间格式弄成excel里的时间格式,
    merge_Excel(0)
    #获得统计sheet
    getAllData(session,0)

    if (int(isHalfYear) != 0):
        # 得到上半年的任务考核数据
        getMonthTasksExcel(isHalfYear)
        getMonthBugsSolveExcel(isHalfYear)
        getMonthBugsReturnExcel(session, isHalfYear)
        getMonthBugsCreateExcel(session,isHalfYear)
        getMonthBugsCloseExcel(session,isHalfYear)
        getMonthTestExcel(isHalfYear)
        merge_Excel(isHalfYear)
        # 获得统计sheet
        # 获得半年的统计
        getAllData(session, isHalfYear)

    # 今天
    today_day=time.strftime("%d")
    # 要每月1号才导出所有需求
    if(today_day=="01") or (today_day=="1") :
        # 1.登录打开网页
        driver = open_browser_login()
        # 得到每个产品线的所有产品需求
        getRequestFiles(driver,session)
        #修改CSV文件
        change_request_csv(session)
        #合并文件
        merge_csv("requests", 1, 0, False)
    # 今天
    today_day = time.strftime("%d")
    # 要每月1号才导出所有工时消耗
    try:
        if (today_day == "01") or (today_day == "1"):
            getMonthTasksConsume(0)
        else:
            # 要每月26号才导出所有工时消耗判断本月的
            getMonthTasksConsume2(0)
    except :
        pass

