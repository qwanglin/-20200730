# -*- coding: utf-8 -*- 

"""
# @Time : 2019/4/11 
# @Author : xucaimin
"""
import json
import xlrd
import datetime
import time
import os
#两个列表合并成新的list
def list_dic(list1,list2):
    '''
    two lists merge a dict,a list as key,other list as value
    :param list1:key
    :param list2:value
    :return:dict
    '''
    dic = list(map(lambda x,y:str([x,y]), list1,list2))
    return dic
def list_dic2(list1,list2):
    '''
    two lists merge a dict,a list as key,other list as value
    :param list1:key
    :param list2:value
    :return:dict
    '''
    dic = dict(map(lambda x,y:[x,y], list1,list2))
    return dic
#判断是否闰年
def is_leap_year(year_num):
    if year_num % 100 == 0:
        if year_num % 400 == 0:
            return True
        else:
            return False
    else:
        if year_num % 4 == 0:
            return True
        else:
            return False


def FIND(string):
    #定义两个变量：分别表示开始的字符串，结束的字符串
    start1= ''
    end1 = '产品'
    #使用find找到开始和结束截取的位置
    s = string.find(start1)
    e = string.find(end1)
    sub_str = string[s:e + len(end1)]
    # print("截取的产品名字：",sub_str)
    return sub_str


#
def getAllDatas(ws1Datas,ws2Datas,ws3Datas,ws4Datas,ws6Datas,ws7Datas,ws10Datas):
    AllDatas=[]
    #上述所有数组都包含默认的3个字段，所有数组里都要加3个字段
    # ws1Datas:5,ws2Datas:7,ws3Datas:3,ws4Datas:4,ws6Datas:2,ws7Datas:3,ws10Datas:1
    #总共有28个字段
    for ws in ws1Datas:
        ad=ws
        for i in range(0, 20):  # 最后加几个0
            ad.append(0)
        AllDatas.append(ad)

    for ws in ws2Datas:
        ad=[]
        for w in range(0,len(ws)):
            if(w<3):
                ad.append(ws[w])
        for i in range(0, 5):  # 最后加几个0
            ad.append(0)
        for w in range(0,len(ws)):
            if(w>2):
                ad.append(ws[w])
        for i in range(0, 13):  # 最后加几个0
            ad.append(0)
        AllDatas.append(ad)

    for ws in ws3Datas:
        ad = []
        for w in range(0, len(ws)):
            if (w < 3):
                ad.append(ws[w])
        for i in range(0, 12):  # 最后加几个0
            ad.append(0)
        for w in range(0, len(ws)):
            if (w > 2):
                ad.append(ws[w])
        for i in range(0, 10):  # 最后加几个0
            ad.append(0)
        AllDatas.append(ad)

    for ws in ws4Datas:
        ad = []
        for w in range(0, len(ws)):
            if (w < 3):
                ad.append(ws[w])
        for i in range(0, 15):  # 最后加几个0
            ad.append(0)
        for w in range(0, len(ws)):
            if (w > 2):
                ad.append(ws[w])
        for i in range(0, 6):  # 最后加几个0
            ad.append(0)
        AllDatas.append(ad)

    for ws in ws6Datas:
        ad = []
        for w in range(0, len(ws)):
            if (w < 3):
                ad.append(ws[w])
        for i in range(0, 19):  # 最后加几个0
            ad.append(0)
        for w in range(0, len(ws)):
            if (w > 2):
                ad.append(ws[w])
        for i in range(0, 4):  # 最后加几个0
            ad.append(0)
        AllDatas.append(ad)

    for ws in ws7Datas:
        ad = []
        for w in range(0, len(ws)):
            if (w < 3):
                ad.append(ws[w])
        for i in range(0,21):  # 最后加几个0
            ad.append(0)
        for w in range(0, len(ws)):
            if (w > 2):
                ad.append(ws[w])
        for i in range(0,1):  # 最后加几个0
            ad.append(0)
        AllDatas.append(ad)

    for ws in ws10Datas:
        ad = []
        for w in range(0, len(ws)):
            if (w < 3):
                ad.append(ws[w])
        for i in range(0,24):  # 最后加几个0
            ad.append(0)
        for w in range(0, len(ws)):
            if (w > 2):
                ad.append(ws[w])
        AllDatas.append(ad)
    print("总统计的数据AllDatas1:",len(AllDatas),AllDatas)

    #把AllDatas数组里的值相同名字的都相加，产品线也叠加起来
    wsAllDatas=[]
    for index1 in range(0, len(AllDatas)-1):
        item1=AllDatas[index1]
        # print("11第几个：",index1,item1[1],len(AllDatas))
        # print("总的多长：", len(wsAllDatas), wsAllDatas)
        for index2 in range(index1+1, len(AllDatas)):
            item2=AllDatas[index2]
            # print("2第几个：",index2,item2[1])
            if ((item1[1] == item2[1])): #名字相同
                # print("名字相同",item1[1], item2[1],index1,index2)
                flag122=False
                flag123 = False
                for index3 in range(0,len(wsAllDatas)):
                    item3=wsAllDatas[index3]
                    if(item3[1]==item1[1]): #总消耗数组里也有该名字
                        flag122=True
                        # print("总产品线：",item2[2],item3[2])
                        if (item2[2] in item3[2]):  # 如果产品线在总的产品线里
                            flag123 = True
                            break
                        break
                # print("总数据中名字产品线相同与否的标志：",flag122,flag123)
                if(flag122==True):#总数据中有
                    # print(item2[2],item3[2])
                    a = item3[2].split('、')
                    b = item2[2].split('、')
                    # print("产品线：",a,b)
                    c = list(set(a).union(set(b)))  # 把产品线合并
                    productli = ""
                    for ii in range(0, len(c)):
                        pro = c[ii]
                        if (pro != ""):
                            if (ii != 0):
                                productli = productli + "、" + pro
                            else:
                                productli = pro
                    item3[2]=productli
                    for iii in range(3,len(item2)):
                        item3[iii] = item3[iii] + item2[iii]
                else:#总数据中也没有
                    wsData=[]
                    wsData.append(item1[0])
                    wsData.append(item2[1])
                    a = item1[2].split('、')
                    b = item2[2].split('、')
                    c = list(set(a).union(set(b)))  # 把产品线合并
                    productli = ""
                    for ii in range(0, len(c)):
                        pro = c[ii]
                        if (pro != ""):
                            if (ii != 0):
                                productli = productli + "、" + pro
                            else:
                                productli = pro
                    wsData.append(productli)
                    for iii in range(3,len(item2)):
                        wsData.append(item1[iii] + item2[iii])
                    wsAllDatas.append(wsData)
                break
    # print("总统计的数据AllDatas2:", len(wsAllDatas), wsAllDatas)
    #还要再做一次筛选，当产品线数量为1的时候删除掉
    wsAllDatas = [j for j in wsAllDatas if len(j[2].split('、')) > 1]
    print("总统计的数据AllDatas3:", len(wsAllDatas), wsAllDatas)
    return wsAllDatas




#判断日期为星期几
def getWeekday(year,mon,day):

    list1 = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]  # 闰年2月份为29天
    list2 = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]  # 平年2月份为28天
    date = 0
    years = 0

    # 输入的年份大于等于2018年的判断过程如下:
    if year >= 2018:
        for j in range(2018, year):
            if (j % 4 == 0) & (j % 100 != 0) or j % 400 == 0:  # 闰年
                years += 366
            else:  # 平年
                years += 365  # 闰年天数加366天,平年加365天

        if ((year % 4) == 0) & ((year % 100) != 0) or ((year % 400) == 0):
            for i in range(mon - 1):
                date += list1[i]  # 闰年月份按list1相加
            days = date + day
        else:
            for i in range(mon - 1):
                date += list2[i]  # 平年月份按list2相加
            days = date + day

        total = days + years
        ji = total % 7  # 参考日期是2018年1月1号是星期一

        # 由于"ji=0"时,输出的结果是"星期0",因此对"ji"进行了判断,使"ji=0"时输出的结果为"星期7"
        if ji != 0:
            weeday=ji
            print(year, '年', mon, '月', day, '日', '是星期', ji)
        else:
            weeday=7
            print(year, '年', mon, '月', day, '日', '是星期', 7)

    # 输入的年份小于2018年的判断过程如下:
    else:
        for j in range(year + 1, 2018):
            if (j % 4 == 0) & (j % 100 != 0) or j % 400 == 0:
                years += 366
            else:
                years += 365

        if ((year % 4) == 0) & ((year % 100) != 0) or ((year % 400) == 0):
            for i in range(mon - 1, 12):
                date += list1[i]
            days = date - day + 1
        else:
            for i in range(mon - 1, 12):
                date += list2[i]
            days = date - day + 1

        total = days + years
        ji = total % 7

        if ji != 0:
            # 余数为1是星期7,余数为2是星期6...,总结规律为8-ji
            weeday=8-ji
            print(year, '年', mon, '月', day, '日', '是星期', 8 - ji)
        else:
            weeday=1
            print(year, '年', mon, '月', day, '日', '是星期', 1)
    print(weeday)



#要操作数据库--查询姓名
def get_user_operate_sql(sqlhost,sqlport,sqluser,sqlpwd,sqldb):
    import pymysql
    userids=[]
    useraccounts=[]
    usernames=[] #姓名中文名
    usernamepys=[] #姓名拼音
    # 数据库连接,括号中依次是你要连接数据库所在的电脑ip, port,用户名，密码，数据库名称,字符集（可加可不加)
    con = pymysql.connect(host=sqlhost, port=sqlport,user=sqluser, passwd=sqlpwd, db=sqldb,charset="utf8")
    cursor = con.cursor()#cursor游标
    selectsql = 'select * from user'
    # selectsql='select user_id from user'
    try:
        cursor.execute(selectsql)  # 执行语句，返回受影响的行数，主要用于执行
        result = cursor.fetchall()  # 获取查询结果
        print("result:",result)
        for row in result:
            userid=row[0]
            useraccount=row[1]
            username=row[2]
            usernamepy=row[7]
            if (str(useraccount) != "" and str(useraccount)!=None):
                userids.append(userid)
                useraccounts.append(useraccount)
                usernames.append(username)  # 中文
                usernamepys.append(usernamepy)  # 拼音

            # print("user_id:",userid)
            # print("row:",row)
    except:
        print("Error: unable to fecth data")
    cursor.close()
    con.close()
    users=[]
    users.append(userids)
    users.append(useraccounts)
    users.append(usernames)
    users.append(usernamepys)
    return users,userids,useraccounts,usernames,usernamepys





#得到相同任务负责人下的任务编号和任务名称
#[[[],[[],[]]],[[],[[],[]]]]
def get_task_bySameRespname(taskDatas):
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




#得到相同bug负责人下的bug编号和bug标题
#[[[],[[],[]]],[[],[[],[]]]]
def get_bug_bySameRespname(bugDatas):
    bugrespNames = []  #bug责任人
    bugcreateNames=[] #bug创建人
    for i in range(0, len(bugDatas)):
        if (i != 0):
            flag1=False
            flag2=False
            bugData = bugDatas[i]
            bugId=bugData[0]
            bugName=bugData[1]
            bugcreateName=bugData[2]
            bugcreateTime=bugData[3]
            bugrespName=bugData[4]
            bugresqsolveTime=bugData[5]
            for j in range(0,len(bugrespNames)):
                bugresp = bugrespNames[j]
                bug1 = bugresp[0]
                bug2 = bugresp[1]
                if (str(bug1[0]) == str(bugrespName)):
                    flag1 = True
                    bug2.append(bugData)
                    break
            if (flag1 == False):
                bugresp = []
                bug1 = []
                bug2 = []
                bug1.append(bugrespName)
                bug2.append(bugData)
                bugresp.append(bug1)
                bugresp.append(bug2)
                bugrespNames.append(bugresp)


            for j in range(0,len(bugcreateNames)):
                bugcreate = bugcreateNames[j]
                bug1 = bugcreate[0]
                bug2 = bugcreate[1]
                if (str(bug1[0]) == str(bugcreateName)):
                    flag2 = True
                    bug2.append(bugData)
                    break
            if (flag2 == False):
                bugcreate = []
                bug1 = []
                bug2 = []
                bug1.append(bugcreateName)
                bug2.append(bugData)
                bugcreate.append(bug1)
                bugcreate.append(bug2)
                bugcreateNames.append(bugcreate)

    print("bugrespNames:",len(bugrespNames),bugrespNames)
    print("bugcreateNames:",len(bugcreateNames),bugcreateNames)
    return bugrespNames,bugcreateNames




#得到相同名字（负责人、创建人）的任务、bug信息
#[[1,2,......]]
def get_taskbug_bySameName(taskrespNames,taskcreateNames,bugrespNames,bugcreateNames):

    taskbugs=[]
    today=time.strftime("%Y/%m/%d")
    print("today:",today)
    #任务负责人
    for i in range(0,len(taskrespNames)):
        taskresp=taskrespNames[i]
        task1=taskresp[0]
        task2=taskresp[1]
        for k in range(0,len(task2)):
            taskbuginfoFlag=""
            taskbuginfo=""   #任务详情
            flag = False
            t2=task2[k]
            taskId = t2[0]
            taskName = t2[1]
            taskexceptstartTime = t2[2]
            taskactuastartTime = t2[3]
            taskendTime = t2[4] #截止时间
            taskcreateName = t2[6]
            taskcreateTime = t2[7]
            taskrespName = t2[8]
            yearnow=int(today.split("/")[0].strip())
            monthnow=int(today.split("/")[1].strip())
            daynow=int(today.split("/")[2].strip())
            today1=datetime.datetime(yearnow,monthnow,daynow) #今天
            # 计算中间时间点，1.开始时间有，那就开始时间，2，无，那就创建时间
            if ("0000" not in str(taskexceptstartTime)):#预计开始时间
                yearstart = int(str(taskexceptstartTime).split("/")[0])
                monthstart = int(str(taskexceptstartTime).split("/")[1])
                daystart = int(str(taskexceptstartTime).split("/")[2])
            elif ("0000" not in str(taskactuastartTime)):#实际开始时间
                yearstart = int(str(taskactuastartTime).split("/")[0])
                monthstart = int(str(taskactuastartTime).split("/")[1])
                daystart = int(str(taskactuastartTime).split("/")[2])
            elif ("0000" not in str(taskcreateTime)):#创建时间
                yearstart = int(str(taskcreateTime).split("/")[0])
                monthstart = int(str(taskcreateTime).split("/")[1])
                daystart = int(str(taskcreateTime).split("/")[2])
            start1 = datetime.datetime(yearstart, monthstart, daystart)  # 任务开始日期
            end_start_2=0  #中间时间点
            end_today_day=0  #截止时间和今天相差几天
            today_start_day=0 #今天和任务开始时间相差几天
            if ("0000" not in str(taskendTime)):  # 如果截止时间不为空
                yearend = int(str(taskendTime).split("/")[0])
                monthend = int(str(taskendTime).split("/")[1])
                dayend = int(str(taskendTime).split("/")[2])
                end1=datetime.datetime(yearend, monthend, dayend) #截止日期
                today_start_day = (today1 - start1).days
                end_today_day = (end1 - today1).days
                end_start_2=int((end1-start1).days/2)
            else:#如果截止时间为空，那就截止一周
                today_start_day = (today1 - start1).days
                end_today_day = 7 - today_start_day
                end_start_2 = int((today_start_day + end_today_day) / 2)
            if (end_today_day < 0):  # 如果小于0，那就是已经延期
                if(str(taskcreateName)==str(taskrespName)):
                    taskbuginfoFlag = "您创建并负责的任务"
                    taskbuginfo = "您创建并负责的任务已延期"+str(
                        abs(end_today_day)) +"天,【" + str(taskId) + "+" + str(taskName) + "】,请尽快完成"
                else:
                    taskbuginfoFlag = "您负责的任务"
                    taskbuginfo = "您负责的任务已延期" + str(
                        abs(end_today_day)) + "天,【" + str(taskId) + "+" + str(taskName) + "】,请尽快完成"
            else:  # 没延期
                if (str(taskcreateName) == str(taskrespName)):
                    if (end_today_day == 0):
                        taskbuginfoFlag="您创建并负责的任务"
                        taskbuginfo = "您创建并负责的任务今天截止,【" + str(taskId) + "+" + str(taskName) + "】，请尽快完成"
                    elif (today_start_day == end_start_2):
                        taskinfoFlag = "您创建并负责的任务"
                        taskbuginfo = "您创建并负责的任务时间已过半,【" + str(taskId) + "+" + str(taskName) + "】，请尽快完成"
                    elif (end_today_day == 2):
                        taskbuginfoFlag = "您创建并负责的任务"
                        taskbuginfo = "您创建并负责的任务时间还剩2天,【" + str(taskId) + "+" + str(taskName) + "】，请尽快完成"
                else:
                    if (end_today_day == 0):
                        taskbuginfoFlag = "您负责的任务"
                        taskbuginfo = "您负责的任务今天截止,【" + str(taskId) + "+" + str(taskName) + "】，请尽快完成"
                    elif (today_start_day == end_start_2):
                        taskbuginfoFlag = "您负责的任务"
                        taskbuginfo = "您负责的任务时间已过半,【" + str(taskId) + "+" + str(taskName) + "】，请尽快完成"
                    elif (end_today_day == 2):
                        taskbuginfoFlag = "您负责的任务"
                        taskbuginfo = "您负责的任务时间还剩2天,【" + str(taskId) + "+" + str(taskName) + "】，请尽快完成"
            for j in range(0, len(taskbugs)):
                tb = taskbugs[j]
                tbname = tb[0]  #负责人/创建人
                if (str(tbname) == str(task1[0])):
                    flag = True
                    if (taskbuginfo != ""):
                        #还要判断是否要排序，如果排序，那就insert,如果不排序，那就append
                        flag2=False
                        index=0
                        for m in range(1,len(tb)):
                            tbinfo=tb[m]
                            if(taskbuginfoFlag in tbinfo):
                                #如果存在就排序对比
                                if(index==0):
                                    index = m
                                flag2=True
                                taskbuginfoFlag2=taskbuginfoFlag+"已延期"
                                if(taskbuginfoFlag2 in tbinfo and taskbuginfoFlag2 in taskbuginfo):
                                    tbinfo_day = tbinfo.split(taskbuginfoFlag2)[1].split("天")[0]
                                    taskbuginfo_day=taskbuginfo.split(taskbuginfoFlag2)[1].split("天")[0]
                                    if(int(taskbuginfo_day)>=int(tbinfo_day)):
                                        index=m
                                        break
                                    else:
                                        index=m+1
                                taskbuginfoFlag2 = taskbuginfoFlag + "还剩"
                                if (taskbuginfoFlag2 in tbinfo and taskbuginfoFlag2 in taskbuginfo):
                                    index=m
                                    break
                                taskbuginfoFlag2 = taskbuginfoFlag + "今天截止"
                                if (taskbuginfoFlag2 in tbinfo and taskbuginfoFlag2 in taskbuginfo):
                                    index=m
                                    break
                                taskbuginfoFlag2 = taskbuginfoFlag + "时间已过半"
                                if (taskbuginfoFlag2 in tbinfo and taskbuginfoFlag2 in taskbuginfo):
                                    index=m
                                    break
                        if(flag2==False):#不排序
                            tb.append("")
                            tb.append(taskbuginfo)
                        else:#如果flag2=true,里面有该关键字
                            tb.insert(index,taskbuginfo)
                    break

            if (flag == False and taskbuginfo!=""):  #当没有该负责人/创建人时，就新建一个数组
                tb = []
                tb.append(taskrespName)
                tb.append(taskbuginfo)
                taskbugs.append(tb)

    #任务创建人
    for i in range(0,len(taskcreateNames)):
        taskcreate=taskcreateNames[i]
        task1=taskcreate[0]
        task2=taskcreate[1]
        for k in range(0,len(task2)):
            taskbuginfoFlag = ""
            taskbuginfo=""   #任务详情
            flag = False
            t2=task2[k]
            taskId = t2[0]
            taskName = t2[1]
            taskexceptstartTime = t2[2]
            taskactuastartTime = t2[3]
            taskendTime = t2[4] #截止时间
            taskcreateName = t2[6]
            taskcreateTime = t2[7]
            taskrespName = t2[8]
            yearnow=int(today.split("/")[0].strip())
            monthnow=int(today.split("/")[1].strip())
            daynow=int(today.split("/")[2].strip())
            today1=datetime.datetime(yearnow,monthnow,daynow) #今天
            end_today_day=0  #截止时间和今天相差几天
            if ("0000" not in str(taskendTime)):  # 如果截止时间不为空
                yearend = int(str(taskendTime).split("/")[0])
                monthend = int(str(taskendTime).split("/")[1])
                dayend = int(str(taskendTime).split("/")[2])
                end1=datetime.datetime(yearend, monthend, dayend) #截止日期
                end_today_day = (end1 - today1).days
                if (end_today_day < 0):  # 如果小于0，那就是已经延期
                    if(str(taskcreateName)!=str(taskrespName)):
                        taskbuginfoFlag="您创建的任务"
                        taskbuginfo = "您创建的任务已延期"+str(
                            abs(end_today_day))+"天,【" + str(taskId) + "+" + str(taskName) + "】,请您关注"
            else:#如果截止时间为空，那就提示还未安排计划
                taskbuginfoFlag="您创建的任务"
                taskbuginfo = "您创建的任务还未排时间计划，【" + str(taskId) +"+"+ str(taskName) + "】，请尽快安排"
            for j in range(0, len(taskbugs)):
                tb = taskbugs[j]
                tbname = tb[0]  #负责人/创建人
                if (str(tbname) == str(task1[0])):
                    flag = True
                    if (taskbuginfo != ""):
                        #还要判断是否要排序，如果排序，那就insert,如果不排序，那就append
                        flag2=False
                        index=0
                        for m in range(1,len(tb)):
                            tbinfo=tb[m]
                            if(taskbuginfoFlag in tbinfo):
                                if (index == 0):
                                    index = m
                                #如果存在就排序对比
                                flag2=True
                                taskbuginfoFlag2=taskbuginfoFlag+"已延期"
                                if(taskbuginfoFlag2 in tbinfo and taskbuginfoFlag2 in taskbuginfo):
                                    tbinfo_day = tbinfo.split(taskbuginfoFlag2)[1].split("天")[0]
                                    taskbuginfo_day=taskbuginfo.split(taskbuginfoFlag2)[1].split("天")[0]
                                    if(int(taskbuginfo_day)>=int(tbinfo_day)):
                                        index=m
                                        break
                                    else:
                                        index=m+1
                                taskbuginfoFlag2 = taskbuginfoFlag + "还未排时间计划"
                                if (taskbuginfoFlag2 in tbinfo and taskbuginfoFlag2 in taskbuginfo):
                                    index=m
                                    break
                        if(flag2==False):#不排序
                            tb.append("")
                            tb.append(taskbuginfo)
                        else:
                            tb.insert(index,taskbuginfo)
                    break
            if (flag == False and taskbuginfo!=""):
                tb = []
                tb.append(taskcreateName)
                tb.append(taskbuginfo)
                taskbugs.append(tb)


    #bug负责人
    for i in range(0,len(bugrespNames)):
        bugresp=bugrespNames[i]
        print("bug解决：",bugresp)
        bug1=bugresp[0]
        bug2=bugresp[1]
        for k in range(0,len(bug2)):
            taskbuginfoFlag=""
            taskbuginfo=""   #bug详情
            flag = False
            b2=bug2[k]
            bugId = b2[0]
            bugName = b2[1]
            bugcreateName = b2[2]
            bugcreateTime = b2[3]
            bugrespName = b2[4]
            bugrequsolveTime = b2[5]#要求解决时间
            yearnow=int(today.split("/")[0].strip())
            monthnow=int(today.split("/")[1].strip())
            daynow=int(today.split("/")[2].strip())
            today1=datetime.datetime(yearnow,monthnow,daynow) #今天
            end_today_day=0  #截止时间和今天相差几天
            if ("0000" not in str(bugrequsolveTime)):  # 如果要求解决时间不为空
                yearend = int(str(bugrequsolveTime).split("/")[0])
                monthend = int(str(bugrequsolveTime).split("/")[1])
                dayend = int(str(bugrequsolveTime).split("/")[2])
                end1=datetime.datetime(yearend, monthend, dayend) #要求解决时间
                end_today_day = (end1 - today1).days
                if (end_today_day < 0):  # 如果小于0，那就是已经延期
                    if(str(bugcreateName)==str(bugrespName)):
                        taskbuginfoFlag="您创建并且需要解决的Bug已延期"
                        taskbuginfo = "您创建并且需要解决的Bug已延期"+str(
                            abs(end_today_day))+"天,【" + str(bugId) + "+" + str(bugName) + "】,请尽快解决"
                    else:
                        taskbuginfoFlag="您需要解决的Bug已延期"
                        taskbuginfo = "您需要解决的Bug已延期"+str(
                            abs(end_today_day))+"天，【" + str(bugId) + "+" + str(bugName) + "】,请尽快解决"

            for j in range(0, len(taskbugs)):
                tb = taskbugs[j]
                tbname = tb[0]  #负责人/创建人
                if (str(tbname) == str(bug1[0])):
                    flag = True
                    if (taskbuginfo != ""):
                        #还要判断是否要排序，如果排序，那就insert,如果不排序，那就append
                        flag2=False
                        index=0
                        for m in range(1,len(tb)):
                            tbinfo=tb[m]
                            if(taskbuginfoFlag in tbinfo):
                                #如果存在就排序对比
                                if (index == 0):
                                    index = m
                                flag2=True
                                tbinfo_day = tbinfo.split(taskbuginfoFlag)[1].split("天")[0]
                                taskbuginfo_day = taskbuginfo.split(taskbuginfoFlag)[1].split("天")[0]
                                if (int(taskbuginfo_day) >= int(tbinfo_day)):
                                    index = m
                                    break
                                else:
                                    index = m + 1

                        if(flag2==False):#不排序
                            tb.append("")
                            tb.append(taskbuginfo)
                        else:
                            tb.insert(index,taskbuginfo)
                    break

            if (flag == False and taskbuginfo!=""):
                tb = []
                tb.append(bugrespName)
                tb.append(taskbuginfo)
                taskbugs.append(tb)


    #bug创建人
    for i in range(0,len(bugcreateNames)):
        bugcreate=bugcreateNames[i]
        bug1=bugcreate[0]
        bug2=bugcreate[1]
        for k in range(0,len(bug2)):
            taskbuginfoFlag=""
            taskbuginfo=""   #bug详情
            flag = False
            b2=bug2[k]
            bugId = b2[0]
            bugName = b2[1]
            bugcreateName = b2[2]
            bugcreateTime = b2[3]
            bugrespName = b2[4]
            bugrequsolveTime = b2[5]#要求解决时间
            yearnow=int(today.split("/")[0].strip())
            monthnow=int(today.split("/")[1].strip())
            daynow=int(today.split("/")[2].strip())
            today1=datetime.datetime(yearnow,monthnow,daynow) #今天
            end_today_day=0  #截止时间和今天相差几天
            if ("0000" not in str(bugrequsolveTime)):  # 如果要求解决时间不为空
                yearend = int(str(bugrequsolveTime).split("/")[0])
                monthend = int(str(bugrequsolveTime).split("/")[1])
                dayend = int(str(bugrequsolveTime).split("/")[2])
                end1=datetime.datetime(yearend, monthend, dayend) #要求解决时间
                end_today_day = (end1 - today1).days
                if (end_today_day < 0):  # 如果小于0，那就是已经延期
                    if (str(bugcreateName) != str(bugrespName)):
                        taskbuginfoFlag="您创建的Bug已延期"
                        taskbuginfo = "您创建的Bug已延期"+str(
                            abs(end_today_day)) +"天，【" + str(bugId) + "+" + str(bugName) + "】,请及时跟踪Bug"

            for j in range(0, len(taskbugs)):
                tb = taskbugs[j]
                tbname = tb[0]  #负责人/创建人
                if (str(tbname) == str(bug1[0])):
                    flag = True
                    if (taskbuginfo != ""):
                        #还要判断是否要排序，如果排序，那就insert,如果不排序，那就append
                        flag2=False
                        index=0
                        for m in range(1,len(tb)):
                            tbinfo=tb[m]
                            if(taskbuginfoFlag in tbinfo):
                                #如果存在就排序对比
                                if (index == 0):
                                    index = m
                                flag2=True
                                tbinfo_day = tbinfo.split(taskbuginfoFlag)[1].split("天")[0]
                                taskbuginfo_day = taskbuginfo.split(taskbuginfoFlag)[1].split("天")[0]
                                if (int(taskbuginfo_day) >= int(tbinfo_day)):
                                    index = m
                                    break
                                else:
                                    index = m + 1
                        if(flag2==False):#不排序
                            tb.append("")
                            tb.append(taskbuginfo)
                        else:
                            tb.insert(index,taskbuginfo)
                    break
            if (flag == False and taskbuginfo!=""):
                tb = []
                tb.append(bugcreateName)
                tb.append(taskbuginfo)
                taskbugs.append(tb)


    print("taskbugs:", len(taskbugs),taskbugs)
    for i in range(0,len(taskbugs)):
        print("iiii:",taskbugs[i][0],taskbugs[i])
    print("taskbugs:",len(taskbugs))
    return taskbugs



#把任务、bug信息转化为图片，并且保存在一个文件夹中
def get_taskbug_image(taskbugs,useraccounts,usernames,usernamepys):
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
    for i in range(0,len(taskbugs)):
        taskbug=taskbugs[i]
        maxLength=0
        for k in range(0, len(taskbug)):
            lentb=len(taskbug[k])
            if(lentb>maxLength):
                maxLength=lentb
        # im = Image.new("RGB", (maxLength*20, len(taskbug) * 36), (255, 255, 255))
        im=PIL.Image.new("RGB", (maxLength * 20, len(taskbug) * 36), (255, 255, 255))
        dr =PIL.ImageDraw.Draw(im)
        #simsun.ttc字体样式在c:/windows/fonts下选择
        font =PIL.ImageFont.truetype(os.path.join("fonts", "simsun.ttc"), 14)
        taskbugname = taskbug[0]
        if(taskbugname==""):
            continue
        taskbugnamepy=""
        for k in range(0,len(usernames)):
            if(str(taskbugname)==str(usernames[k])):
                taskbugnameaccount=useraccounts[k]
                if(str(taskbugnameaccount)!="" and str(taskbugnameaccount)!=None):
                    taskbugnamepy=useraccounts[k]
                else:
                    taskbugnamepy = usernamepys[k]
                break
        if(taskbugnamepy!=""):
            for j in range(0, len(taskbug)):
                tb = taskbug[j]
                if (j == 0):
                    tb2 = str(now)+" 禅道任务bug状态推送 "+tb + ":"
                else:
                    if(str(tb)==""):
                        tb2 = ""
                    else:
                        tb2 = tb + ";"
                dr.text((3, 2+(j * 28)), str(tb2), font=font, fill="#000000")
            # im.show()
            # print("taskbugname:",taskbugname)
            im.save(imagespath + "/" + str(taskbugnamepy) + ".png")


if __name__=="__main__":
    # print(is_leap_year(2008))
    # indexs=[3,4,2,3,5,1,0,2,3,5]
    # a="jjjo哈哈" \
    #   "哈"
    # d1 = datetime.datetime(2019, 11, 1)
    # d2 = datetime.datetime(2019, 10, 31)
    # print((d1-d2).days)
    # # text=[['徐采敏','您负责的任务【4395+任务过程状态推送/企业微信接口研究】已延期2天,请尽快完成','您创建的任务【4253+GTRB-S单板支持基于IMSI向用户发送短信】还未排时间计划，请尽快安排',"您创建的任务【3963+客户端链路根据热点工作频点和SIM卡的运营商类型，自动确定工作模式和频段】还未排时间计划，请尽快安排"]]
    # # get_taskbug_image(text)
    #
    # # path = os.path.dirname(os.getcwd()) + '/Images'
    # # from common .getfileName import *
    # # Imagespath = new_file(path)  # 得到结果文件最新的文件夹
    # # print(Imagespath)
    # # users,userids, usernames, usernamepys = get_user_operate_sql("192.168.9.247", 3310, "pm", "pm@zed!", "eoffice10")
    # # print("usernames:",usernames)
    # num_list=["你好1","你好10","你好3"]
    # num_list.sort(reverse=True)
    # print(num_list)
    # a=[1,2]
    # a.insert(1,444)
    # print(a)
    import PIL.Image,PIL.ImageFont,PIL.ImageDraw
    now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    imagespath = os.path.dirname(os.getcwd()) + "/Images/" + datetime.datetime.now().strftime('%Y%m%d %H%M%S')
    print("imagespath:",imagespath)
    try:
        os.makedirs(imagespath)
    except:
        pass
    im = PIL.Image.new("RGB", (20 * 30, 10*20), (255, 255, 255))
    dr = PIL.ImageDraw.Draw(im)
    # simsun.ttc字体样式在c:/windows/fonts下选择
    font = PIL.ImageFont.truetype(os.path.join("fonts", "simsun.ttc"), 14)
    str1 = str(now) + " 禅道截至当前当月工时消耗状态推送  " +"顾承浩" + ":" +"\n"*3
    str2 ="""您本月消耗的工时为 0 !!
请检查登录禅道检查您本月是否安排了任务; 
若安排了任务请检查是否开启;
          """
    dr.text((3, 30), str(str1 + str2), font=font, fill="#DC143C")
    im.save(imagespath + "/" + str("guchenghao") + ".png")
