# -*- coding: utf-8 -*- 

"""
# @Time : 2019/3/27 
# @Author : xucaimin
"""

from common.configPar import *
from common.WriteToTxt import logWriteToTxt
import requests
import json
import urllib3
import re
import urllib
from bs4 import BeautifulSoup
from urllib import request



#####产品线-》产品--》迭代--》任务


url=getConfig("DEFAULT","url")

headers= {
        "Connection": "keep-alive",
        "Cache-Control": "max-age=0",
        "Content-Type": "application/x-www-form-urlencoded",
        "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.86 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        "Referer": "http://192.168.10.237/zentao/user-login.html",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9",
        }

#接口账号登录
def _login():
    loginurl=url+"/zentao/user-login.html"
    print("loginurl:",loginurl)
    session = requests.session()
    session.verify = False
    account=getConfig("DEFAULT","account")
    password=getConfig("DEFAULT","password")
    body = {"account": account,  # 你自己的账号
            "password": password,  # 你自己的密码
            "keepLogin[]": "on",
            "referer":"/zentao/my/"
            }
    rs = session.post(url=loginurl, data=body)
    rs.encoding = 'utf-8'
    # print(rs.text)
    return session

#_getiterate给这个方法提供链接
def _project(session):

    #得到迭代主页的下拉框任务里的链接
    rs = session.get(url + "/zentao/project-index-no.html")
    rs.encoding = 'utf-8'
    # print(rs.text)
    soup = BeautifulSoup(rs.text, "html.parser")
    projexturl=""
    for tag in soup.find_all('div', id='subHeader'):
        # print("2222222222",tag)
        for item in tag.find_all('div', class_='dropdown-menu'):
            # print("3333",item)
            # print("444",item.attrs)
            # print("555",item.attrs.get("data-url"))
            projexturl=item.attrs.get("data-url")
    print(projexturl+"\n")
    return projexturl

#获取所有的迭代（已关闭的迭代）
def _getAllClosediterate(session):
    closedUrls=[]
    closedTimes=[]
    rs = session.get(url + "/zentao/project-all-closed-51-order_desc-0.html") #找到已关闭的迭代
    rs.encoding = 'utf-8'
    sou = BeautifulSoup(rs.text, "html.parser")
    for ul in sou.find_all('ul', class_='pager'):
        pager = ul.attrs.get('data-rec-total')
        # print("pager:",pager)
        # upager = i.split('.')[0] + "-order_desc-" + pager + "-2000-1.html"
        # upagers.append(upager)
    iteurl=url+"/zentao/project-all-closed-51-order_desc-0-"+pager+"-2000-1.html "
    rs2=session.get(iteurl)
    soup = BeautifulSoup(rs2.text, "html.parser")
    for tr in soup.find_all('tr'):
        for td in tr.find_all('td',class_='text-left'):
            if ("title" in str(td)):
                for a in td.find_all('a'):
                    href = a.attrs.get('href')
                    closedUrls.append(href)
    # print("closedUrls:",len(closedUrls),closedUrls)
    for closedUrl in closedUrls:
        rs = session.get(url+closedUrl)
        # print("1:",closedUrl)
        rs.encoding = 'utf-8'
        sou = BeautifulSoup(rs.text, "html.parser")
        t="无"
        for ol in sou.find_all("ol",class_='histories-list'):
            for li in ol.find_all('li'):
                if("关闭" in str(li)):
                    strli=str(li)
                    # print("2:",closedUrl,strli.split(',')[0].split('>')[1].strip())
                    t=strli.split(',')[0].split('>')[1].strip().split(' ')[0]
                    # print(t)
                    break
        closedTimes.append(t)
    print("closedTimes:",len(closedTimes),closedTimes)
    return closedTimes



#获取所有未关闭的迭代
def _getiterate(session):

    projexturl=_project(session)
    rs = session.get(url + projexturl)
    rs.encoding = 'utf-8'
    # print(rs.text)
    global  gloabhrefs
    global gloabamounts
    global gloabtitles
    hrefs=""
    amounts = ""
    titles = ""
    soup = BeautifulSoup(rs.text, "html.parser")
    for tag in soup.find_all('div', class_='table-col col-left'):
        # print("222",tag)
        for item in tag.find_all('div', class_='list-group'):
            # print("333",item)
            a = item.findAll('a')
            # print("444",a)
            # print(len(a))

            for ai in a:
                # print(ai)
                href=ai.attrs.get("href")
                title=ai.attrs.get("title")
                amount = re.sub('\D', '', href)
                hrefs = hrefs + href
                hrefs = hrefs + ","
                amounts = amounts + amount
                amounts = amounts + ","
                titles = titles + title
                titles = titles + ","
                # print(href)
                # print(amount)
                # print(title)
            # print(hrefs)
            # print(2,amounts)
            # print(titles)
            gloabhrefs=hrefs
            gloabamounts=amounts
            gloabtitles=titles
    print("gloabtitles,gloabamounts:",gloabtitles,gloabamounts)
    return gloabamounts


#得到迭代-bug的id
def _getIteraBugId(session):
    amounts = _getiterate(session)
    a = amounts.split(',')
    hrefs=[]
    dates=[]
    for i in a:
        if (i != ""):
            buildurl = url + "/zentao/project-build-" + i + ".html"
            rs = session.get(buildurl)
            rs.encoding = 'utf-8'
            # print(rs.text)
            soup = BeautifulSoup(rs.text, "html.parser")
            for tag in soup.find_all('td', class_='c-name'):
                # print(tag)
                for a in tag.find_all('a'):
                    # print(a)
                    href=a.attrs.get('href')
                    title=a.string
                    hrefs.append(href)
                    # print(href,title)
            for tag in soup.find_all('td',class_='c-date'):
                date=tag.string
                dates.append(date)
    # print(len(hrefs),hrefs)
    # print(len(dates),dates)
    dateids=[]  #版本日期与bug的id的集合
    for i in range(0,len(hrefs)):
        buildviewurl = url +hrefs[i]
        rs = session.get(buildviewurl)
        rs.encoding = 'utf-8'
        # print(rs.text)
        soup = BeautifulSoup(rs.text, "html.parser")
        for tag in soup.find_all('td', class_='c-id text-left'):
            # print(tag)
            for div in tag.find_all('div',class_='checkbox-primary'):
                # print(div)
                for label in div.find_all('label') :
                    # print(label)
                    bugid=label.string
                    dateid=dates[i]+"+"+bugid
                    # print(bugid,dateid)
                    dateids.append(dateid)
    print(len(dateids),dateids)
    return dateids




#获取bug的最后一次激活的时间
def _getBugActivityLastTime(session,bugid):
    bugurl=url+"/zentao/bug-view-"+str(bugid)+".html"
    rs = session.get(bugurl)
    rs.encoding = 'utf-8'
    # print(rs.text)
    soup = BeautifulSoup(rs.text, "html.parser")
    firsttime=""
    for tag in soup.find_all('ol',class_='histories-list'):
        # print(tag)
        for li in tag.find_all('li'):
            # print(1111,str(li))
            if("激活" in str(li)):
                # print(22222, li)
                # print(222,str(li).split(',')[0].split('>')[1].lstrip().split(" ")[0])
                #第一次激活时间
                # firsttime=str(li).split(',')[0].split('>')[1].lstrip().split(" ")[0]
                # ftime=str(firsttime.split('-')[0]) + "/" + str(firsttime.split('-')[1]) + "/" + str(firsttime.split('-')[2])
                # break
                #最后一次激活时间
                lasttime = str(li).split(',')[0].split('>')[1].lstrip().split(" ")[0]
                ltime = str(lasttime.split('-')[0]) + "/" + str(lasttime.split('-')[1]) + "/" + str(lasttime.split('-')[2])
    return ltime

#获取所有任务
def _getTasks(session):
    amounts=_getiterate(session)
    allproducts=_getProduct(session)
    products=allproducts[1]
    # print("products:",products)
    allproducttasks=[]
    a = amounts.split(',')
    for i in a:
        if(i!=""):
            taskurl=url+"/zentao/project-task-"+i+"-all.html"
            # print("taskurl:",taskurl)
            rs = session.get(taskurl)
            rs.encoding = 'utf-8'
            soup = BeautifulSoup(rs.text, "html.parser")
            for tag in soup.find_all('span', class_='label label-light label-badge'):
                amount=tag.string
                break
            taskurl = url + "/zentao/project-task-" + i + "-all-0--"+amount+"-2000-1.html"
            # print("taskurl2:",taskurl)
            rs = session.get(taskurl)
            rs.encoding = 'utf-8'
            # print(rs.text)
            soup = BeautifulSoup(rs.text, "html.parser")
            for tag in soup.find_all('div', class_='main-col'):
                # print(2,tag)
                for item in tag.find_all('td', class_='c-name'):
                    # print(item)
                    for a in item.findAll('a'):
                        # print(a)
                        href = a.attrs.get("href")
                        if(href!=None):
                            taskname = a.string.strip()
                            taskurl2 = url + href
                            # print("taskurl2:", taskurl2)
                            rs2 = session.get(taskurl2)
                            rs2.encoding = 'utf-8'
                            soup2 = BeautifulSoup(rs2.text, "html.parser")
                            for tag2 in soup2.find_all('div', id='legendBasic'):
                                # print(tag2)
                                flag = False
                                tr2 = ""
                                for tr in tag2.findAll('tr'):
                                    for th in tr.findAll('th'):
                                        if (th.string == "所属模块"):
                                            flag = True
                                            tr2 = tr
                                            break
                                # print(tr2)
                                for td in tr2.findAll('td'):
                                    # print(td)
                                    title=td.attrs.get('title')
                                    # print(title)
                            if(title!="/"):
                                if("/" in title):
                                    title=title.split('/')[0]
                                for pro in products:
                                    for p in pro:
                                        if(p==title):
                                            title=title.strip()
                                            protask=title+"+"+taskname
                                            allproducttasks.append(protask)
                                            break
    print("allproducttasks:",allproducttasks)
    return allproducttasks


##得到产品线下的任务（有需求的任务）
#得到产品线下的产品
def _getProduct(session):
    rs = session.get(url+"/zentao/product-all.html")
    rs.encoding = 'utf-8'
    # print(rs.text)
    soup = BeautifulSoup(rs.text, "html.parser")
    hrefs = ""
    products=[]
    for tag in soup.find_all('ul', id='modules'):
        # print(tag)
        a = tag.findAll('a')
        # print(a)
        for item in a:
            # print(item)
            href = item.attrs.get("href")
            # print(href)
            hrefs = hrefs + href
            hrefs = hrefs + ","
            product = item.string
            # print(product)
            products .append(product)

    allIteras = []  # 产品线所有的迭代allIteras=[]  #产品线所有的迭代
    h = hrefs.split(',')
    allPros = []  # 产品线所有的产品
    upagers=[] #链接
    for i in h:
        if i != "":
            Iteras = []  # 一个产品线所有的迭代
            Pros = []  # 一个产品线下所有的产品名称
            r = session.get(url + i)
            r.encoding = 'utf-8'
            sou = BeautifulSoup(r.text, "html.parser")
            for ul in sou.find_all('ul',class_='pager'):
                pager=ul.attrs.get('data-rec-total')
                upager=i.split('.')[0]+"-order_desc-"+pager+"-2000-1.html"
                upagers.append(upager)

    allPros=[]  #产品线所有的产品
    for i in upagers:
        Iteras = []  # 一个产品线所有的迭代
        Pros = []  # 一个产品线下所有的产品名称
        r = session.get(url + i)
        r.encoding = 'utf-8'
        # print(i,r.text)
        # print(i)
        sou = BeautifulSoup(r.text, "html.parser")
        producthrefs = ""
        for tag in sou.find_all('div', class_='main-col'):
            # print(i,tag)
            for item in tag.find_all('td', class_='c-name'):
                # print(item)
                for a in item.findAll('a'):
                    href = a.attrs.get("href")
                    producthrefs = producthrefs + href
                    producthrefs = producthrefs + ","
                    pro = a.string
                    Pros.append(pro)
                    # print(href,a.string)
        allPros.append(Pros)


    '''
            ph = producthrefs.split(',')
            for p in ph:
                if p!="":
                    r2 = session.get(url + p)
                    r2.encoding = 'utf-8'
                    # print(p)
                    sou = BeautifulSoup(r2.text, "html.parser")
                    for tag in sou.find_all('div', id='subHeader'):
                        # print(tag)
                        for item in tag.find_all('a'):
                            # print(item)
                            if(item.string=="迭代"):
                                allproducthref = item.attrs.get("href")
                                # print(allproducthref)
                                r3 = session.get(url + allproducthref)
                                r3.encoding = 'utf-8'
                                # print(rs.text)
                                so = BeautifulSoup(r3.text, "html.parser")
                                for tag in so.find_all('td',class_='text-left'):
                                    # print(tag)
                                    for t in tag.find_all('a'):
                                        text = t.string
                                        # print(text)
                                        Iteras.append(text)
            new_Iteras = list(set(Iteras))#list数组去重
            # print(new_Iteras)
            allIteras.append(new_Iteras)
    '''
    all=[]
    # print(products)
    # print(allIteras)
    all.append(products)#产品线
    all.append(allPros)#产品
    # all.append(allIteras)#所有产品线下的所有产品的所有迭代(迭代下有任务）
    print("all:",all)
    return all

#下载文件
def _download(session):
    #1.获取
    downurl=url+"/zentao/task-export-88-status,id_desc-unclosed.html"
    # print(downurl)
    names = gloabtitles.split(',')
    # print(names)
    # fileName=names[2]+"-未关闭任务"
    fileName="ZEDNE Sprint020-未关闭任务"
    # print(fileName)
    headers = {'Content-Type': 'application/json'}  # 设置数据为json格式，很重要
    body = {
           "title": "默认模板",
            "template": 0,
            "fileType": "csv",
            "fileName": fileName,
            "encode":"UTF-8",
            # "exportType":all,
            "exportFields[]":["id","project","module","story","name"]
            # "exportFields[]":["id","project","module","story","name","files","type","pri","estStarted","realStarted","deadline","status","desc","left",
            #                   "lastEditedDate","finishedDate","progress","openedBy","openedDate","assignedTo","consumed","finishedBy","canceledBy","closedReason",
            #                   "canceledDate","closedBy","closedDate","estimate","mailto","lastEditedBy","assignedDate"],
            }
    rs = session.post(url=downurl,data=body)
    rs.encoding = 'utf-8'
    # print(rs.text)
    # print("333",rs.content)
    # fp = open("yoyos.xlsx", "wb")
    # fp.write(rs.content)
    # fp.close()

#得到部门下的人
def _getDepartment(session):
    companyurl=url+"/zentao/company-browse.html"
    rs = session.get(companyurl)
    rs.encoding = 'utf-8'
    # print(rs.text)
    soup = BeautifulSoup(rs.text, "html.parser")
    hrefs=[]
    departments=[]
    for tag in soup.find_all('ul', class_='tree'):
        # print(2,tag)
        for li in tag.find_all('li'):
            # print(1,li)
            for a in li.find_all('a'):
                # print(2,a.attrs.get("href"),a.string)
                hrefs.append(a.attrs.get("href"))
                departments.append(a.string)
    # new_hrefs = list(set(hrefs))  # list数组去重
    # new_departments=list(set(departments))# list数组去重,但是会改变顺序
    #接下来去重不会改变顺序
    new_hrefs = []
    new_departments = []
    for i in hrefs:
        if i not in new_hrefs:
            new_hrefs.append(i)
    for i in departments:
        if i not in new_departments:
            new_departments.append(i)
    # print(new_hrefs)
    # print(new_departments)
    new_hrefs2 = []
    for i in range(0, len(new_departments)):
        # print(new_departments[i])
        # print(url + new_hrefs[i])
        rs = session.get(url + new_hrefs[i])
        rs.encoding = 'utf-8'
        soup = BeautifulSoup(rs.text, "html.parser")
        for tag in soup.find_all('ul', class_='pager'):
            # print(1,tag)
            mount = tag.attrs.get('data-rec-total')
            # print(mount)
            newhref = new_hrefs[i].split('.')[0] + "-bydept-id-" + mount + "-2000-1.html"  # 设置每页2000
            new_hrefs2.append(newhref)
    # print(new_hrefs2)

    peoples = []
    for i in range(0, len(new_departments)):
        people = []
        # print(new_departments[i])
        if (new_departments[i] != "研发体系"):
            # print(url+new_hrefs2[i])
            rs = session.get(url + new_hrefs2[i])
            rs.encoding = 'utf-8'
            # print(rs.text)
            soup = BeautifulSoup(rs.text, "html.parser")
            # print("i:",i)
            for tag in soup.find_all('tbody'):
                # print(i,tag)
                for td in tag.find_all('td'):
                    # print(td)
                    for a in td.find_all('a'):
                        # print(1,a)
                        if ((a.string != None) & (a.attrs.get('title') != None)):
                            # print(a.string,a.attrs.get('title'))
                            people.append(a.string)
        peoples.append(people)
    # print(peoples)
    depatementpeoples=[]
    depatementpeoples.append(new_departments)
    depatementpeoples.append(peoples)
    print("depatementpeoples:",depatementpeoples)
    return depatementpeoples


#测试-测试单-版本-用例
#得到版本情况
def _getVersions(session):
    qaurl = url + "/zentao/qa/"
    rs = session.get(qaurl)
    rs.encoding = 'utf-8'
    # print(rs.text)
    soup = BeautifulSoup(rs.text, "html.parser")
    for tag in soup.find_all('nav', id='subNavbar'):
        # print(tag)
        for li in tag.find_all('li'):
            # print(li)
            for a in li.find_all('a'):
                if(a.string=="测试单"):
                    href = a.attrs.get('href')
                    amount = re.sub('\D', '', href)
    dropurl=url+"/zentao/product-ajaxGetDropMenu-"+str(amount)+"-testtask-browse-.html"
    # print("默认下拉框dropurl:",dropurl)
    hrefs=[]
    rs = session.get(dropurl)#得到测试下拉框下主页所有的链接
    rs.encoding = 'utf-8'
    # print(rs.text)
    soup = BeautifulSoup(rs.text, "html.parser")
    for tag in soup.find_all('div', class_='table-col col-left'):
        for div in tag.find_all('div',class_='list-group'):
            for a in div.find_all('a'):
                hrefs.append(a.attrs.get('href'))
    # print("下拉框hrefs:",hrefs)
    testhrefs=[]
    for h in hrefs:
        rs = session.get(url+h)
        rs.encoding = 'utf-8'
        # print(rs.text)
        if('table-footer' in rs.text):
            soup = BeautifulSoup(rs.text, "html.parser")
            for tag in soup.find_all('div', class_='table-footer'):
                for ul in tag.find_all('ul'):
                    testmount=ul.attrs.get('data-rec-total')
                    testhrefs.append(url+h.split('.')[0]+"--local,totalStatus-id_desc-"+str(testmount)+"-2000-1-0-0.html")
    print("下拉框并且选择2000testhrefs:",testhrefs)
    return testhrefs



#测试-测试单-版本-用例
#得到版本下的用例情况
def _getTestCases(session):
    versionurls = []#版本链接
    versions=[]  #版本号
    testhrefs=_getVersions(session)
    for testhref in testhrefs:
        rs = session.get(testhref)
        rs.encoding = 'utf-8'
        # print(rs.text)
        soup = BeautifulSoup(rs.text, "html.parser")
        for tbody in soup.find_all('tbody'):
            for tr in tbody.find_all('tr'):
                versionurl = ""
                for a in tr.find_all('a'):
                    if (versionurl != ""):
                        break
                    versionurl = a.attrs.get('href')
                    # print("versionurl:",versionurl)
                    versionurls.append(versionurl)
            for tr in tbody.find_all('tr'):
                id = ""
                versionname = ""
                productname = ""
                iteraname = ""
                versionnum = ""
                people = ""

                version=[]
                start=""
                end=""
                state=""
                for td in tr.find_all('td'):
                    if (id == ""):
                        for a in td.find_all('a'):
                            id = a.string
                            break
                    elif (versionname == ""):
                        for a in td.find_all('a'):
                            versionname = a.string
                            break
                    elif (productname == ""):
                        productname = td.string
                    elif (iteraname == ""):
                        iteraname = td.string
                    elif (versionnum == ""):
                        if ("/zentao/build-view" in str(td)):
                            for a in td.find_all('a'):
                                versionnum = a.string
                                break
                        else:
                            versionnum = td.string
                    elif (people == ""):
                        people = td.string
                    elif (start == ""):
                        start = td.string
                    elif (end == ""):
                        end = td.string
                    elif (state == ""):
                        state = td.attrs.get('title')
                    else:
                        break
                version.append(versionnum)
                version.append(start)
                version.append(end)
                version.append(state)
                version.append(productname)
                versions.append(version)

    print("versionurls:", len(versionurls), versionurls)
    print("versions:", len(versions), versions)
    versionurls2=[]
    for v in versionurls:
        rs = session.get(url+v)
        rs.encoding = 'utf-8'
        # print(rs.text)
        if ("data-rec-total" in str(rs.text)):
            soup = BeautifulSoup(rs.text, "html.parser")
            for ul in soup.find_all('ul', class_='pager'):
                total = ul.attrs.get('data-rec-total')
                versionurls2.append(url + v.split('.')[0] + "-all-0-id_desc-" + total + "-2000-1.html")
        else:
            versionurls2.append(url+v)
    print("versionurls2:",len(versionurls2),versionurls2)
    all=[]
    all.append(versions)
    all.append(versionurls2)
    print("all:",all)
    return all

#判断日期是否为节假日或周末
def getHoliday(day):
    url = "http://api.goseek.cn/Tools/holiday?date="+day
    rs=requests.get(url)
    # print(day,rs.text)
    info_dict = json.loads(rs.text)
    # print(info_dict.get("data"))
    return info_dict.get("data")

#得到多人任务下的责任人任务的完成情况
def getManyResponseTask(session,index):
    rs = session.get(url + "/zentao/task-view-"+str(index)+".html")
    rs.encoding = 'utf-8'
    # print(rs.text)
    NameTimes=[]
    Names=[]
    Times=[]
    soup = BeautifulSoup(rs.text, "html.parser")
    for ol in soup.find_all('ol', class_='histories-list'):
        # print("111:",ol)
        for li in ol.findAll("li"):
            # print("222:",li)
            if ("由" in str(li) and "完成" in str(li)):
                finishtime=str(li).strip().split(',')[0].strip().split('>')[1].strip()
                finishname=str(li).strip().split(',')[1].strip().split('</')[0].strip().split('>')[1].strip()
                # print("333:",finishname,finishtime)
                flag=True
                for na in Names:
                    if(finishname ==na):
                        flag=False
                        break
                if(flag==True):
                    Names.append(finishname)
                    Times.append(finishtime)
    NameTimes.append(Names)
    NameTimes.append(Times)
    print("NameTimes:",NameTimes)
    return NameTimes


if __name__=="__main__":

    # session=_login()
    # _project(session)
    import time
    today = time.strftime("%Y%m%d %H%M%S")  # 今天
    print("执行时间为：" + today + " " )





















