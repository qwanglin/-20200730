# -*- coding: utf-8 -*-
import sys



sys.path.append("../")
from common.configPar import *
from selenium import webdriver
import time
from common.WriteToTxt import *
from common.zentaoInterface import _login,_getiterate,_getProduct,_getTasks,_getDepartment,_getIteraBugId,_getAllClosediterate,\
    _getBugActivityLastTime,_getVersions,_getTestCases,getHoliday,getManyResponseTask
from selenium.webdriver.common.action_chains import ActionChains
import os
from datetime import datetime
from common.zentao_login import *
import csv
from common.getfileName import *
import pandas as pd
import glob
from common.csv__xlsx import *
from xlwt import *
import xlrd
from xlutils.copy import copy
from common.commonUtil import *
import openpyxl
import  datetime as dt
from bs4 import BeautifulSoup
from common.seleniumUtil import *
from case.zentao_getTaskBugFiles import *
from case.zentao_notice import WeChat_Email
from common.ftputil import ftp_upload_data_image
from common.get_developers import *
url=getConfig("DEFAULT","url")

# 获取进行中、已完成、已关闭的任务+当月工时数组
def getWorkTimeTask():
    path = os.path.dirname(os.getcwd()) + '/result'
    Folder_Path = new_file(path)  # 得到结果文件最新的文件夹
    # os.chdir(Folder_Path)
    # file_list = os.listdir()
    allpath = Folder_Path + "/工时消耗.xls"
    print("allpath:", allpath)
    rb = xlrd.open_workbook(allpath, formatting_info=True)  # 打开 上月版本测试完成情况.xls文件，把其他文件的值复制到该文件
    ws1 = rb.sheet_by_name("data")  # 获取表单
    nrows1 = ws1.nrows  # 获取行数
    ncols1 = ws1.ncols  # 获取列数
    ws1Datas = []  # 存储任务tasks数据
    for i in range(0, nrows1):
        if(i!=0):
            idcol=0  #编号列
            taskstatus=12# 任务状态
            fathertaskcol=4 #任务名称列
            sontaskcol=5  #子任务名称列
            exceptstartcol=9   #预计开始时间列
            actuastartcol=10   #实际开始时间列
            endtimecol=11 #截止时间列
            taskstatecol=12 #任务状态列
            createnamecol=18  #创建人列
            createtimecol=19 #创建时间列
            namecol = 34 # 责任人列
            mothworktime=38 #当月消耗工时
            taskstatusresult=ws1.cell_value(i,taskstatus).strip()
            if taskstatusresult != "未开始":
                idcolresult=ws1.cell_value(i, idcol).strip()
                fathertaskcolresult = ws1.cell_value(i, fathertaskcol).strip()
                sontaskcolresult = ws1.cell_value(i, sontaskcol).strip()
                exceptstartcolresult = ws1.cell_value(i, exceptstartcol).strip()
                actuastartcolresult = ws1.cell_value(i, actuastartcol).strip()
                endtimecolresult = ws1.cell_value(i, endtimecol).strip()
                taskstatecolresult = ws1.cell_value(i, taskstatecol).strip()
                createnamecolresult = ws1.cell_value(i, createnamecol).strip()
                createtimecolresult = ws1.cell_value(i, createtimecol).strip()
                namecolresult = ws1.cell_value(i, namecol).strip()
                mothworktimeresult=ws1.cell_value(i,mothworktime).strip()
                ws1Data = []
                ws1Data.append(idcolresult)
                if(str(sontaskcolresult)!=""):
                    ws1Data.append(str(fathertaskcolresult) + "/"+str(sontaskcolresult))
                else:
                    ws1Data.append(str(fathertaskcolresult))
                ws1Data.append(exceptstartcolresult)
                ws1Data.append(actuastartcolresult)
                ws1Data.append(endtimecolresult)
                ws1Data.append(taskstatecolresult)
                ws1Data.append(createnamecolresult)
                ws1Data.append(createtimecolresult)
                ws1Data.append(namecolresult)
                ws1Data.append(mothworktimeresult)
                ws1Datas.append(ws1Data)
        else:
            ws1Data=[]
            #9个
            ws1Data.append("任务编号")
            ws1Data.append("任务名称")
            ws1Data.append("预计开始时间")
            ws1Data.append("实际开始时间")
            ws1Data.append("截止日期")
            ws1Data.append("任务状态")
            ws1Data.append("任务创建人")
            ws1Data.append("任务创建时间")
            ws1Data.append("任务责任人")
            ws1Data.append("当月消耗工时")
            ws1Datas.append(ws1Data)

    print("已完成或进行中tasks ws1Datas:",len(ws1Datas),ws1Datas)
    return ws1Datas


#得到相同任务负责人下的任务编号和任务名称
#[[[],[[],[]]],[[],[[],[]]]]
def get_worktime_task_bySameRespname(taskDatas):
    taskrespNames=[]  #任务负责人
    taskcreateNames=[] #任务创建人
    for i in range(0,len(taskDatas)):
        if(i!=0):
            flag1=False
            flag2=False
            taskData=taskDatas[i]
            taskId=taskData[0]
            taskName=taskData[1]
            taskexceptstartTime=taskData[2]
            taskactuastartTime=taskData[3]
            taskendTime=taskData[4]
            taskcreateName=taskData[6]
            taskcreateTime=taskData[7]
            taskrespName=taskData[8]
            for j  in range(0,len(taskrespNames)):
                taskresp=taskrespNames[j]
                task1=taskresp[0]
                task2=taskresp[1]
                if(str(task1[0])==str(taskrespName)):
                    flag1=True
                    task2.append(taskData)
                    break
            if (flag1 == False):
                taskresp = []
                task1 = []
                task2 = []
                task1.append(taskrespName)
                task2.append(taskData)
                taskresp.append(task1)
                taskresp.append(task2)
                taskrespNames.append(taskresp)

            for j  in range(0,len(taskcreateNames)):
                taskcreate=taskcreateNames[j]
                task1 = taskcreate[0]
                task2 = taskcreate[1]
                if (str(task1[0]) == str(taskcreateName)):
                    flag2=True
                    task2.append(taskData)
                    break
            if (flag2 == False):
                taskcreate = []
                task1 = []
                task2 = []
                task1.append(taskcreateName)
                task2.append(taskData)
                taskcreate.append(task1)
                taskcreate.append(task2)
                taskcreateNames.append(taskcreate)

    print("taskrespNames:",len(taskrespNames),taskrespNames)
    print("taskcreateNames:",len(taskcreateNames),taskcreateNames)
    return  taskrespNames,taskcreateNames

def get_worktime_tasks(tasks):

    taskbugs=[]
    for task in tasks:
        t = []
        name=task[0][0]
        # print("name:"+str(name))
        content=task[1]
        # print("content:"+str(content))
        t.append(name)
        # contents=[]
        for con in content:
            # print(con)
            newContent="任务编号："+con[0]+" 任务：【"+con[1] +"】 状态："+con[5] +" 截至当前本月消耗总工时："+con[-1]+" 请及时确认自己工时"
            t.append(newContent)
        # t.append(contents)
        taskbugs.append(t)

    print(taskbugs)
    return taskbugs

#把任务、bug信息转化为图片，并且保存在一个文件夹中
def get_worktime_task_image_new(taskbugs,useraccounts,usernames,usernamepys):
    # import Image, ImageFont, ImageDraw
    # import pygame


    import PIL.Image,PIL.ImageFont,PIL.ImageDraw
    now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    imagespath = os.path.dirname(os.getcwd()) + "/Images/" + datetime.datetime.now().strftime('%Y%m%d %H%M%S')
    print("imagespath:",imagespath)
    try:
        os.makedirs(imagespath)
    except:
        pass
    tasknames=[]
    for index,taskbug  in enumerate(taskbugs):
        taskbugname = taskbug[0]
        tasknames.append(taskbugname)
        maxLength=0
        for info  in  taskbug:
            lentb=len(info)
            if (lentb>maxLength):
                maxLength=lentb
        im = PIL.Image.new("RGB", (maxLength * 20, len(taskbug) * 36), (255, 255, 255))
        dr =PIL.ImageDraw.Draw(im)
        #simsun.ttc字体样式在c:/windows/fonts下选择
        font =PIL.ImageFont.truetype(os.path.join("fonts", "simsun.ttc"), 14)

        if (taskbugname == ""):
            continue
        taskbugnamepy = ""
        for k,username in enumerate(usernames):
            if (str(username) == str(taskbugname)):
                taskbugnameaccount = useraccounts[k]
                if (str(taskbugnameaccount) != "" and str(taskbugnameaccount) != None):
                    taskbugnamepy = useraccounts[k]
                else:
                    taskbugnamepy = usernamepys[k]
                break

        if(taskbugnamepy!=""):
            for j,tb in enumerate(taskbug):

                if (j == 0):
                    tb2 = str(now)+" 禅道截至当前当月工时消耗状态推送  "+tb + ":"
                else:
                    if(str(tb)==""):
                        tb2 = ""
                    else:
                        tb2 = tb + ";"
            # tb2 = str(now)+" 禅道当月工时消耗推送 "+taskbugname + ":"+"\n"
            # contents=taskbug[1]
            # for index,content in enumerate(contents):
            #     tb2=tb2+content+";\n"
                dr.text((3, 2+(j * 28)), str(tb2), font=font, fill="#000000")
            # im.show()
            print("taskbugname:",taskbugname)
            im.save(imagespath + "/" + str(taskbugnamepy) + ".png")
    item_users=[]
    for index,username in enumerate(usernames):
        if username  not in tasknames:
            item_users.append(username)
            im = PIL.Image.new("RGB", (20 * 30, 10 * 20), (255, 255, 255))
            dr = PIL.ImageDraw.Draw(im)
            # simsun.ttc字体样式在c:/windows/fonts下选择
            font = PIL.ImageFont.truetype(os.path.join("fonts", "simsun.ttc"), 14)
            str1 = str(now) + " 禅道截至当前当月工时消耗状态推送  " + username + ":" + "\n" * 3
            str2 = """您本月消耗的工时为 0 !!
请登录禅道检查您本月是否安排了任务; 
若安排了任务请检查是否开启;
                      """
            taskbugnameaccount2 = useraccounts[index]
            if (str(taskbugnameaccount2) != "" and str(taskbugnameaccount2) != None):
                taskbugnamepy2 = useraccounts[index]
            else:
                taskbugnamepy2 = usernamepys[index]
            print("不在的用户为：" + str(username))
            dr.text((3, 30), str(str1 + str2), font=font, fill="#0000FF")
            im.save(imagespath + "/" + str(taskbugnamepy2) + ".png")


    return item_users



class Logger(object):
  def __init__(self, filename="Default.log"):
    self.terminal = sys.stdout
    self.log = open(filename, "a")
  def write(self, message):
    self.terminal.write(message)
    self.log.write(message)
  def flush(self):
    pass


if __name__ == '__main__':
    fpath = os.path.split(sys.path[0])[0]
    sys.stdout = Logger(fpath + '/logger.txt')
    today = time.strftime("%Y%m%d %H%M%S")  # 今天
    print("执行时间为：" + today + " " + fpath)
    isHalfYear = getConfig("DEFAULT", "isHalfYear")
    # 1.登录打开网页
    driver = open_browser_login()
    # 2.得到迭代的文件--下载到本地（指定目录)
    session = getTaskFiles(driver, isHalfYear)
    # print("session:"+session)
    # 3.修改csv文件
    change_task_csv(session)
    # 4.合并csv文件并转化为excel
    taskxlsxpath = merge_csv("tasks", session, 0, False)
    # 第1个sheet：前一周已完成的任务和所有未完成的任务、未开始和进行中
    getWeekTasksExcel()  # 得到上一周的任务数据（已完成的任务和所有未完成的任务、未开始和进行中--截止时间小于）
    # 第2个sheet：前一月已完成的任务和所有未完成的任务
    today_day = time.strftime("%d")
    getMonthTasksConsume2(0)
    d = getWorkTimeTask()
    # for i in d :
    #     print(i)
    print("*********************")
    taskrespNames, taskcreateNames = get_worktime_task_bySameRespname(d)
    worktimetask = get_worktime_tasks(taskrespNames)
    sqlhost = str(getConfig("mass", "sqlhost"))
    sqlport = int(getConfig("mass", "sqlport"))
    sqluser = str(getConfig("mass", "sqluser"))
    sqlpwd = str(getConfig("mass", "sqlpwd"))
    sqldb = str(getConfig("mass", "sqldb"))
    users, userids, useraccounts, usernames, usernamepys = get_user_operate_sql(sqlhost, sqlport, sqluser,
                                                                                sqlpwd,
                                                                                sqldb)
    itemusers=get_worktime_task_image_new(worktimetask, useraccounts, usernames, usernamepys)
    ftp_upload_data_image()
    wx = WeChat_Email()
    # users1 = [["WV00000342","WV00000182"], ["guchenghao","zhangchunjuan"], ["顾承浩","张春娟"],
    #           ["guchenghao","zhangchunjuan"]]
    wx.send_text2(worktimetask, users)
    zed=Zendao()
    dev_users=list(zed.get_name_list().keys())
    wx.send_text_notTask(itemusers,users,dev_users)

    # """
    # WV00000182,zhangchunjuan,张春娟,zhangchunjuan
    # """

