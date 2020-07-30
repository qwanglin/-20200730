#-*-coding:utf-8-*-
import datetime
from email.mime.text import MIMEText
from email.header import Header
from email.utils import parseaddr, formataddr
import smtplib
import os
import ftplib
from time import ctime
from email.mime.multipart import MIMEMultipart
import time

def listdir(path, list_name):  #传入存储的list
    for file in os.listdir(path):
        file_path = os.path.join(path, file)
        if os.path.isdir(file_path):
            listdir(file_path, list_name)
        else:
             list_name.append((file_path,os.path.getctime(file_path)))

def newestfile(target_list):
    newest_file = target_list[0]
    for i in range(len(target_list)):
        if i < (len(target_list)-1) and newest_file[1] < target_list[i+1][1]:
            newest_file = target_list[i+1]
        else:
            continue
    # print('newest file is',newest_file)
    return newest_file

def getNewestFilepath(path):
    list=[]
    listdir(path,list)
    newest_file=newestfile(list)
    newest_file_path=newest_file[0]
    return newest_file_path



def ftp_upload_data():
    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    date = datetime.datetime.now().day
    print(year, month, date)
    if month < 10:
        dirMonth = str(year) + "0" + str(month)
    else:
        dirMonth = str(year) + str(month)
    if month<10:
        if date>=10:
            dirDate = "0"+str(month) + "-" + str(date)
        else:
            dirDate = "0" + str(month) + "-0" + str(date)
    else:
        if date>=10:
            dirDate = str(month) + "-" + str(date)
        else:
            dirDate = str(month) + "-0" + str(date)
    ftp = ftplib.FTP('192.168.9.201', user='platform', passwd='zed@pd2018')
    try:
        ftp.cwd("/ZEDNE/tools/zedat/testreport/zentao/")
        ftp.mkd(str(year))
    except:
        print("已经有这个目录")
    try:
        ftp.cwd("/ZEDNE/tools/zedat/testreport/zentao/"+str(year)+"/")
        ftp.mkd(dirMonth)
    except:
        print("已经有这个目录")
    try:
        ftp.cwd("/ZEDNE/tools/zedat/testreport/zentao/"+str(year)+"/" + dirMonth + "/")
        ftp.mkd(dirDate)
    except:
        print("已经有这个目录")

    ftp.cwd("/ZEDNE/tools/zedat/testreport/zentao/"+str(year)+"/" + dirMonth + "/"+dirDate)

    path = os.path.dirname(os.getcwd()) + '/result'
    Folder_Path = os.path.join(path,new_file(path))  # 得到结果文件最新的文件夹,然后上传这个文件夹
    try:
        ftp_upload_dir(Folder_Path,ftp,target_dir="/ZEDNE/tools/zedat/testreport/zentao/"+str(year)+"/" + dirMonth + "/"+dirDate)
    except:
        pass


#上传文件夹（文件夹下有图片）
def ftp_upload_data_image():
    year = datetime.datetime.now().year
    month = datetime.datetime.now().month
    date = datetime.datetime.now().day
    print(year, month, date)
    if month < 10:
        dirMonth = str(year) + "0" + str(month)
    else:
        dirMonth = str(year) + str(month)
    if month<10:
        if date>=10:
            dirDate = "0"+str(month) + "-" + str(date)
        else:
            dirDate = "0" + str(month) + "-0" + str(date)
    else:
        if date>=10:
            dirDate = str(month) + "-" + str(date)
        else:
            dirDate = str(month) + "-0" + str(date)
    ftp = ftplib.FTP('192.168.9.201', user='platform', passwd='zed@pd2018')
    try:
        ftp.cwd("/ZEDNE/tools/zedat/testreport/zentao/")
        ftp.mkd(str(year))
    except:
        print("已经有这个目录")
    try:
        ftp.cwd("/ZEDNE/tools/zedat/testreport/zentao/"+str(year)+"/")
        ftp.mkd(dirMonth)
    except:
        print("已经有这个目录")
    try:
        ftp.cwd("/ZEDNE/tools/zedat/testreport/zentao/"+str(year)+"/" + dirMonth + "/")
        ftp.mkd(dirDate)
    except:
        print("已经有这个目录")

    path = os.path.dirname(os.getcwd()) + '/Images'
    new_filename=new_file(path)  #最新文件夹
    Folder_Path =os.path.join(path,new_filename)   # 得到图片文件最新的文件夹,然后上传这个文件夹
    try:
        ftp.cwd("/ZEDNE/tools/zedat/testreport/zentao/"+str(year)+"/" + dirMonth + "/"+dirDate+"/")
        ftp.mkd(str(new_filename))
    except:
        print("已经有这个目录")
    ftp.cwd("/ZEDNE/tools/zedat/testreport/zentao/" + str(year) + "/" + dirMonth + "/" + dirDate+"/"+str(new_filename))
    try:
        ftp_upload_dir(Folder_Path,ftp,target_dir="/ZEDNE/tools/zedat/testreport/zentao/"+str(year)+"/" + dirMonth + "/"+dirDate+"/"+str(new_filename))
    except:
        pass


#获取文件夹下的最新文件夹或者文件
def new_file(test_file):
    lists = os.listdir(test_file)         # 列出目录的下所有文件和文件夹保存到lists
    lists.sort(key=lambda fn: os.path.getmtime(test_file + "/" + fn)) # 按时间排序
    file_new = os.path.join(test_file, lists[-1])      # 获取最新的文件保存到file_new
    # print("1,",file_new,lists[-1])
    return lists[-1]
    # return file_new


def uploadfile(ftp, remotepath, localpath):
    bufsize = 1024
    fp = open(localpath, 'rb')
    ftp.storbinary('STOR ' + remotepath, fp, bufsize)
    ftp.set_debuglevel(0)
    fp.close()

def ftp_upload_dir(path_source, session, target_dir=None):
    files = os.listdir(path_source)

    # 先记住之前在哪个工作目录中
    last_dir = os.path.abspath('.')
    # 然后切换到目标工作目录
    os.chdir(path_source)

    if target_dir:
        current_dir = session.pwd()
        try:
            session.mkd(target_dir)
        except Exception:
            pass
        finally:
            session.cwd(os.path.join(current_dir, target_dir))

    for file_name in files:
        current_dir = session.pwd()
        if os.path.isfile(path_source + r'/{}'.format(file_name)):
            upload_file(path_source, file_name, session)
        elif os.path.isdir(path_source + r'/{}'.format(file_name)):

            current_dir = session.pwd()
            try:
                session.mkd(file_name)
            except:
                pass
            session.cwd("%s/%s" % (current_dir, file_name))
            ftp_upload_dir(path_source + r'/{}'.format(file_name), session)

        # 之前路径可能已经变更，需要再回复到之前的路径里
        session.cwd(current_dir)

    os.chdir(last_dir)


def upload_file(path, file_name, session, target_dir=None, callback=None):
    # 记录当前 ftp 路径
    cur_dir = session.pwd()

    if target_dir:
        try:
            session.mkd(target_dir)
        except:
            pass
        finally:
            session.cwd(os.path.join(cur_dir, target_dir))

    # print("path:%s \r\n\t   file_name:%s" % (path, file_name))
    file = open(os.path.join(path, file_name), 'rb')  # file to send
    session.storbinary('STOR %s' % file_name, file, callback = callback)  # send the file
    file.close()  # close file
    session.cwd(cur_dir)

def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), addr))

def sendEmail():
    msg = MIMEMultipart()
    from_addr = 'pm@zed.com'
    smtp_server = '192.168.9.247'
    msg['From'] = _format_addr('pm@zed.com')
    receivers = ['guchenghao@zed.com','qiaolin@zed.com','zhangchunjuan@zed.com']
    # receivers = ['youjiabin@zed.com']
    msg['To'] = _format_addr(receivers)
    now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    msg['Subject'] = Header(now+'禅道自动化测试执行完成', 'utf-8').encode()

    mail_msg = '***这是禅道自动化测试，可以进入站点192.168.9.201 /ZEDNE/tools/zedat/testreport/zentao目录下查看文件，不用回复***</p>'
    msg.attach(MIMEText(mail_msg, 'html', 'utf-8'))
    #发送文件
    path = os.path.dirname(os.getcwd())
    print(path)
    Folder_Path = os.path.join(path,new_file(path))
    print(Folder_Path)
    print(Folder_Path+"/工时消耗.xls")

    att = MIMEText(open(Folder_Path+"/all.xls", 'rb').read(), 'base64', 'utf-8')
    att["Content-Type"] = 'application/octet-stream'
    # 这里的filename可以任意写，写什么名字，邮件中显示什么名字
    att["Content-Disposition"] ="attachment;filename=all.xls"
    msg.attach(att)
    if os.path.exists(Folder_Path+"/工时消耗.xls"):
        att2=MIMEText(open(Folder_Path+"/工时消耗.xls", 'rb').read(), 'base64', 'utf-8')
        att2["Content-Type"] = 'application/octet-stream'
        # 这里的filename可以任意写，写什么名字，邮件中显示什么名字
        att2["Content-Disposition"] ="attachment;filename=worktime.xls"
        msg.attach(att2)

    server = smtplib.SMTP(smtp_server, 25, timeout=30)

    server.sendmail('pm@zed.com', ['guchenghao@zed.com','qiaolin@zed.com','zhangchunjuan@zed.com'],
                    msg=msg.as_string())
    server.quit()
    os.system(r'echo %s 邮件发送成功' % ctime())

def sendEmail2():
    msg = MIMEMultipart()
    from_addr = 'guchenghao@zed.com'
    smtp_server = '192.168.9.247'
    msg['From'] = _format_addr('guchenghao@zed.com')
    receivers = ['guchenghao@zed.com','qiaolin@zed.com']
    # receivers = ['youjiabin@zed.com']
    msg['To'] = _format_addr(receivers)
    now = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    msg['Subject'] = Header(now+'禅道自动化测试执行完成', 'utf-8').encode()

    mail_msg = '***这是禅道自动化测试，可以进入站点192.168.9.201 /ZEDNE/tools/zedat/testreport/zentao目录下查看文件，不用回复***</p>'
    msg.attach(MIMEText(mail_msg, 'html', 'utf-8'))
    #发送文件
    path = os.path.dirname(os.getcwd())
    print(path)
    Folder_Path = os.path.join(path,new_file(path))
    print(Folder_Path)
    print(Folder_Path+"/工时消耗.xls")

    att = MIMEText(open(Folder_Path+"/all.xls", 'rb').read(), 'base64', 'utf-8')
    att["Content-Type"] = 'application/octet-stream'
    # 这里的filename可以任意写，写什么名字，邮件中显示什么名字
    att.add_header("Content-Disposition", "attachment", filename=("gbk", "", "all.xls"))
    msg.attach(att)
    if os.path.exists(Folder_Path+"/工时消耗.xls"):
        att=MIMEText(open(Folder_Path+"/工时消耗.xls", 'rb').read(), 'base64', 'utf-8')
        att["Content-Type"] = 'application/octet-stream'
        att.add_header("Content-Disposition", "attachment", filename=("gbk", "", "工时消耗.xls"))
        # 这里的filename可以任意写，写什么名字，邮件中显示什么名字
        # att["Content-Disposition"] = 'attachment; filename="工时消耗.xls"'
        msg.attach(att)

    server = smtplib.SMTP(smtp_server, 25, timeout=30)

    server.sendmail('guchenghao@zed.com', ['guchenghao@zed.com','qiaolin@zed.com'],
                    msg=msg.as_string())
    server.quit()
    os.system(r'echo %s 邮件发送成功' % ctime())


if __name__=="__main__":
    ftp_upload_data()
    sendEmail()


