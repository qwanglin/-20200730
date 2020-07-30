# -*- coding: utf-8 -*- 

"""
# @Author : xucaimin
"""

#发邮件或者企业微信通知相关人员

import sys
sys.path.append("../")
from common.configPar import *
from selenium import webdriver
import time
from common.zentaoInterface import _login
from selenium.webdriver.common.action_chains import ActionChains
import os
import datetime
from common.zentao_login import *
from case.zentao_getTaskBugFiles import *
import requests
import json
import configparser
from common.commonUtil import *
from email.mime.text import MIMEText
from email.header import Header
from email.utils import parseaddr, formataddr
import smtplib
import ftplib
from time import ctime
from email.mime.multipart import MIMEMultipart
from common.configPar import getConfig
from common.ftputil import *

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))


class Logger(object):
  def __init__(self, filename="Default.log"):
    self.terminal = sys.stdout
    self.log = open(filename, "a")
  def write(self, message):
    self.terminal.write(message)
    self.log.write(message)
  def flush(self):
    pass

class Config(object):
    '''解析配置文件'''
    def get_config(self, lable, value):
        cf = configparser.ConfigParser()
        configpath = os.path.dirname(os.getcwd()) + '/config.conf'
        cf.read(configpath,encoding="utf-8-sig")
        config_value = cf.get(lable, value)
        # print("config_value:",config_value)
        return config_value


'''发送信息到企业微信/邮箱'''
class WeChat_Email(Config):

    def __init__(self):
        '''初始化配置'''
        super(Config, self).__init__()
        self.CORPID = self.get_config("mass", "CORPID")
        self.CORPSECRET = self.get_config("mass", "CORPSECRET")
        self.AGENTID = self.get_config("mass", "AGENTID")
        self.TOUSER = self.get_config("mass", "TOUSER")
        self.tokenpath = os.path.dirname(os.getcwd()) + '/access_token.conf'


    #发请求获取token
    def _get_access_token(self):
        '''发起请求'''
        url = 'https://qyapi.weixin.qq.com/cgi-bin/gettoken'
        values = {'corpid': self.CORPID,
                  'corpsecret': self.CORPSECRET,
                  }
        req = requests.post(url, params=values)
        data = json.loads(req.text)
        print ("得到token:",data)
        return data["access_token"]

    #获取token,保存到本地
    def get_access_token(self):
        '''获取token，保存到本地'''
        try:#当有token的时候，就获取token
            with open(self.tokenpath, 'r') as f:
                t, access_token = f.read().split()
        except Exception: #没有就发送请求
            with open(self.tokenpath, 'w') as f:
                access_token = self._get_access_token()
                cur_time = time.time()
                f.write('\t'.join([str(cur_time), access_token]))
                return access_token
        else:#如果没有发生异常，则执行这段代码
            cur_time = time.time()
            if 0 < cur_time - float(t) < 7260:
                return access_token
            else:
                with open(self.tokenpath, 'w') as f:
                    access_token = self._get_access_token()
                    f.write('\t'.join([str(cur_time), access_token]))
                    return access_token

    # 上传到临时素材  图片ID,imagepath不能有中文
    def get_media_ID(self,imagepath):
        Gtoken = self.get_access_token()
        img_url = "https://qyapi.weixin.qq.com/cgi-bin/media/upload?access_token={}&type=image".format(Gtoken)
        files = {'image': open(imagepath, 'rb')}
        r = requests.post(img_url, files=files)
        re = json.loads(r.text)
        # print("media_id:",re['media_id'])
        return re['media_id']


    '''企业微信发送图片'''
    def send_image_WeChat(self,imagepath,touser):  ##发送图片
        import urllib3,urllib
        img_id = self.get_media_ID(imagepath)
        send_values={
            "touser" : touser,
            "msgtype" : "image",
            "agentid" : self.AGENTID,
            "image" : {
                "media_id" : img_id
            },
            "safe":0,
            "enable_duplicate_check": 0,
        }
        send_url= 'https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=' + self.get_access_token()
        # send_msges = (bytes(json.dumps(send_values), 'utf-8'))
        json_send_values=json.dumps(send_values)
        r = requests.post(send_url, json_send_values.encode(encoding='UTF8'))
        print("发送图片返回的值:", r.content)
        return r.content



    '''企业微信发送text信息'''
    def send_text_WeChat(self, taskbugInfo,touser):
        # msg = ''
        # for i in range(0, len(taskbugInfo)):
        #     if (i == 0):
        #         taskbug_username = taskbugInfo[i]
        #         msg = str(taskbug_username) + ":" + "\r"
        #     else:
        #         msg = str(msg) + str(taskbugInfo[i]) +"；"+ "\r"
        send_url = 'https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=' + self.get_access_token()
        send_values = {
            # "touser": self.TOUSER,
            "touser" : touser,
            "msgtype": "text",
            "agentid": self.AGENTID,
            "text": {
                "content": taskbugInfo
                # "content": msg
            },
            "safe": "0"
        }
        #微信内容发送字数有限制??这个该咋解决？发送图片？？？
        send_msges = (bytes(json.dumps(send_values), 'utf-8'))
        r = requests.post(send_url, send_msges)
        print ("发送text返回的值:",r.content)
        return r.content



    '''邮箱发送text信息'''
    def send_text_email(self,taskbugInfo,emailuser):

        msg = MIMEMultipart()
        from_addr = 'pm@zed.com'
        smtp_server = '192.168.9.247'
        msg['From'] = _format_addr('pm@zed.com')
        receivers = [emailuser]
        # receivers = ['youjiabin@zed.com']
        msg['To'] = _format_addr(receivers)
        now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        msg['Subject'] = Header(now + ' 禅道任务bug状态推送', 'utf-8').encode()
        mail_msg='禅道任务bug状态推送,请您关注:'
        for i in range(0,len(taskbugInfo)):
            if(i==0):
                taskbug_username=taskbugInfo[i]
                mail_msg=str(mail_msg)+"</p></p>"
                mail_msg = str(mail_msg)+str(taskbug_username) +":" + "</p></p>"
            else:
                if(str(taskbugInfo[i])!=""):
                    mail_msg = str(mail_msg) + str(taskbugInfo[i]) + "</p>"
                else:
                    mail_msg = str(mail_msg)+ "</p>"

        # mail_msg = '***这是禅道自动化测试，可以进入站点192.168.9.201 /ZEDNE/tools/zedat/testreport/zentao目录下查看文件，不用回复***</p>ttttttt'
        msg.attach(MIMEText(mail_msg, 'html', 'utf-8'))
        # print("要发送的邮件内容:",str(msg))
        try:
            server = smtplib.SMTP(smtp_server, 25, timeout=30)
            server.sendmail('pm@zed.com', [emailuser],
                            msg=msg.as_string())
            server.quit()
            os.system(r'echo %s 邮件发送成功' % ctime())
        except:
            os.system(r'echo %s 邮件发送失败' % ctime())



    '''用企业微信、邮箱发送消息/图片'''
    def send_text(self,taskbugInfos,users):
        now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        userids=users[0]
        useraccounts=users[1]
        usernames=users[2]
        usernamepys=users[3]
        #1.发送消息 2.发送图片
        path = os.path.dirname(os.getcwd()) + '/Images'
        Imagespath = os.path.join(path,new_file(path))  # 得到结果文件最新的文件夹
        print("最新的Images文件夹路径:", Imagespath)
        for i in range(0,len(taskbugInfos)):
            taskbugInfo=taskbugInfos[i]
            # print("iiiii:",i,len(taskbugInfo)-1,taskbugInfo)
            taskbug_username=taskbugInfo[0]  #任务/bug负责人/创建人姓名
            flag=False
            index=-1
            for j in range(0,len(usernames)):
                if(usernames[j]==str(taskbug_username)):
                    flag=True
                    index=j
                    break
            if(flag==True):#姓名存在于OA数据库中，就发送邮件
                if(str(useraccounts[index])!="" and str(useraccounts[index])!=None):
                    usernameaccountpy=useraccounts[index]
                    emailuser = str(useraccounts[index]) + "@zed.com"
                else:
                    usernameaccountpy=usernamepys[index] #拼音
                    emailuser = str(usernamepys[index]) + "@zed.com"
                touser=str(userids[index])
                print("要发送消息给:" + str(usernames[index]))
                print(usernames[index], usernameaccountpy,emailuser, touser)
                try:
                    self.send_text_email(taskbugInfo, emailuser)
                    print("邮件发送成功！")
                except:
                    print("邮件发送失败!")
                time.sleep(2)
                # self.send_text_WeChat("禅道任务bug状态推送,请您关注如下图所示:",touser)
                time.sleep(2)
                try:
                    self.send_image_WeChat(Imagespath + "/" + str(usernameaccountpy) + ".png", touser)
                    print("图片发送成功！")
                except:
                    print("图片发送失败！")
                print("发送成功！！！" + str(usernameaccountpy) + ".png")

        #最后发个消息推送给负责人，告诉负责人推送完成
        self.send_text_WeChat(str(now)+" 禅道任务bug状态全部推送完成", self.TOUSER)

    def send_text2(self, taskbugInfos, users):
        now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        userids = users[0]
        useraccounts = users[1]
        usernames = users[2]
        usernamepys = users[3]
        # 1.发送消息 2.发送图片
        path = os.path.dirname(os.getcwd()) + '/Images'
        Imagespath = os.path.join(path, new_file(path))  # 得到结果文件最新的文件夹
        print("最新的Images文件夹路径:", Imagespath)
        for i in range(0, len(taskbugInfos)):
            taskbugInfo = taskbugInfos[i]
            # print("iiiii:",i,len(taskbugInfo)-1,taskbugInfo)
            taskbug_username = taskbugInfo[0]  # 任务/bug负责人/创建人姓名
            flag = False
            index = -1
            for j in range(0, len(usernames)):
                if (usernames[j] == str(taskbug_username)) and (usernames[j] not in ["段仕勇","王斌","陈工羽","李首忠","罗军","高世洪","张春娟"]):
                    flag = True
                    index = j
                    break
            if (flag == True):  # 姓名存在于OA数据库中，就发送邮件
                if (str(useraccounts[index]) != "" and str(useraccounts[index]) != None):
                    usernameaccountpy = useraccounts[index]
                    emailuser = str(useraccounts[index]) + "@zed.com"
                else:
                    usernameaccountpy = usernamepys[index]  # 拼音
                    emailuser = str(usernamepys[index]) + "@zed.com"
                touser = str(userids[index])
                print("要发送消息给:" + str(usernames[index]))
                print(usernames[index], usernameaccountpy, emailuser, touser)
                try:
                    self.send_text_email(taskbugInfo, emailuser)
                    print("邮件发送成功！")
                except:
                    print("邮件发送失败!")
                time.sleep(2)
                # self.send_text_WeChat("禅道任务bug状态推送,请您关注如下图所示:",touser)
                time.sleep(2)
                try:
                    self.send_image_WeChat(Imagespath + "/" + str(usernameaccountpy) + ".png", touser)
                    print("图片发送成功！")
                except:
                    print("图片发送失败！")
                print("发送成功！！！" + str(usernameaccountpy) + ".png")

        # 最后发个消息推送给负责人，告诉负责人推送完成
        self.send_text_WeChat(str(now) + " 禅道任务bug状态全部推送完成", self.TOUSER)

    def send_text_email_notTask(self,emailuser):
        msg = MIMEMultipart()
        from_addr = 'pm@zed.com'
        smtp_server = '192.168.9.247'
        msg['From'] = _format_addr('pm@zed.com')
        receivers = [emailuser]
        # receivers = ['youjiabin@zed.com']
        msg['To'] = _format_addr(receivers)
        now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        msg['Subject'] = Header(now + ' 禅道任务bug状态推送', 'utf-8').encode()
        mail_msg = '禅道截至当前当月工时消耗状态推送,请您关注:'
        mail_msg= mail_msg+"</p></p>"
        mail_msg=mail_msg+"您本月消耗的工时为 0 !!"+"</p></p>"
        mail_msg=mail_msg+"请登录禅道检查您本月是否安排了任务; "+"</p></p>"
        mail_msg=mail_msg+"若安排了任务请检查是否开启;"+"</p></p>"

        # mail_msg = '***这是禅道自动化测试，可以进入站点192.168.9.201 /ZEDNE/tools/zedat/testreport/zentao目录下查看文件，不用回复***</p>ttttttt'
        msg.attach(MIMEText(mail_msg, 'html', 'utf-8'))
        # print("要发送的邮件内容:",str(msg))
        try:
            server = smtplib.SMTP(smtp_server, 25, timeout=30)
            server.sendmail('pm@zed.com', [emailuser],
                            msg=msg.as_string())
            server.quit()
            os.system(r'echo %s 邮件发送成功' % ctime())
        except:
            os.system(r'echo %s 邮件发送失败' % ctime())
    def send_text_notTask(self,itemusers,users,dev_users):
        """
        发送给没有任务的人
        :param taskbugInfos:
        :param users:
        :param dev_users:研发人员
        :return:
        """
        now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
        userids = users[0]
        useraccounts = users[1]
        usernames = users[2]
        usernamepys = users[3]
        # 1.发送消息 2.发送图片
        path = os.path.dirname(os.getcwd()) + '/Images'
        Imagespath = os.path.join(path, new_file(path))  # 得到结果文件最新的文件夹
        print("最新的Images文件夹路径:", Imagespath)
        for index,username in enumerate(usernames):
           if username in itemusers:
              if  username not in ["段仕勇", "王斌", "陈工羽", "李首忠", "罗军", "高世洪", "张春娟"]:
                   if (str(useraccounts[index]) != "" and str(useraccounts[index]) != None):
                       usernameaccountpy = useraccounts[index]
                       emailuser = str(useraccounts[index]) + "@zed.com"
                   else:
                       usernameaccountpy = usernamepys[index]  # 拼音
                       emailuser = str(usernamepys[index]) + "@zed.com"
                   touser = str(userids[index])
                   print("要发送消息给:" + str(usernames[index]))
                   print(usernames[index], usernameaccountpy, emailuser, touser)
                   if usernameaccountpy in dev_users:
                       try:
                            self.send_text_email_notTask(emailuser)
                            print("邮件发送成功！")
                       except:
                            print("邮件发送失败!")
                       time.sleep(2)
                       # self.send_text_WeChat("禅道任务bug状态推送,请您关注如下图所示:",touser)
                       # time.sleep(2)
                       try:
                           self.send_image_WeChat(Imagespath + "/" + str(usernameaccountpy) + ".png", touser)
                           print("图片发送成功！")
                       except:
                           print("图片发送失败！")
                       print("发送成功！！！" + str(usernameaccountpy) + ".png")

        # 最后发个消息推送给负责人，告诉负责人推送完成
        self.send_text_WeChat(str(now) + " 禅道任务bug状态全部推送完成", self.TOUSER)

if __name__ == '__main__':

    sqlhost = str(getConfig("mass", "sqlhost"))
    sqlport = int(getConfig("mass", "sqlport"))
    sqluser = str(getConfig("mass", "sqluser"))
    sqlpwd = str(getConfig("mass", "sqlpwd"))
    sqldb = str(getConfig("mass", "sqldb"))

    fpath = os.path.split(sys.path[0])[0]
    sys.stdout = Logger(fpath + '/pushlog.txt')  #任务bug推送日志
    today = time.strftime("%Y%m%d %H%M%S")  # 今天
    print("任务过程状态推送执行时间为：" + today + " " + fpath)

    #得到任务task
    # 1.登录打开网页
    driver = open_browser_login()
    # 2.得到迭代的文件--下载到本地（指定目录)
    session = getTaskFiles(driver,0)   #driver页面关闭
    # 3.修改csv文件
    change_task_csv(session)
    # 4.合并csv文件并转化为excel
    taskxlsxpath = merge_csv("tasks", session, 0,True)
    #5.得到状态为未开始和进行中的任务
    getNoStartTasksExcel()


    #得到bug
    # 1.登录打开网页
    driver = open_browser_login()
    # 2.得到bug文件--下载到本地（指定目录
    getBugFiles(driver)
    # 3.修改csv文件
    change_bug_csv(session)
    # 4.合并csv文件，并转化为excel
    merge_csv("bugs", 1, 0,True)
    #5.得到状态为激活的Bug
    getActiviBugExcel()

    #得到未开始或进行中状态的任务数组
    taskDatas=getNoStartTasksData()
    # 得到激活状态的bug数组
    bugDatas=getActiviBugsData()

    # 得到相同任务负责人下的任务编号和任务名称
    taskrespNames, taskcreateNames=get_task_bySameRespname(taskDatas)
    # 得到相同bug负责人下的bug编号和bug标题
    bugrespNames,bugcreateNames=get_bug_bySameRespname(bugDatas)
    #得到名字相同（负责人、创建人）的任务、bug情况
    taskbugs=get_taskbug_bySameName(taskrespNames,taskcreateNames,bugrespNames,bugcreateNames)

    #得到数据库里的姓名和拼音等关键字
    users, userids, useraccounts, usernames, usernamepys =get_user_operate_sql(sqlhost,sqlport,sqluser,sqlpwd,sqldb)

    #把任务、bug信息转化为图片
    get_taskbug_image(taskbugs,useraccounts,usernames,usernamepys)
    ftp_upload_data_image() #把图片上传至ftp

    #最后一步：发送消息
    wx = WeChat_Email()
    wx.send_text(taskbugs,users)
    # 每周三 / 周五禅道任务bug状态推送, 请您关注: < / p > < / p > 单嗣荣: < / p > < / p > 您负责的任务已延期91天,【3995 + 15
    # 系统工程版本 / 生产验证】, 请尽快完成 < / p > 您负责的任务已延期9天,【4704 + 故障（寰创） / 将设备扫频模块换成扫频比较快的模块】, 请尽快完成 < / p >




