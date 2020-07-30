#-*- coding:utf-8 -*-

import requests
from bs4 import BeautifulSoup
import re
import math
from common.WriteToTxt import *
class Zendao(object):
    def __init__(self):
        """
        初始化一个Logger生成器，并生成登录禅道的session
        """

        self.session=requests.session()
        indexPage = self.session.get(url="http://192.168.10.237/zentao/user-login.html")
        logWriteToTxt("访问禅道登录界面")
        verifyRand = self.get_verifyRand(indexPage.text)
        logWriteToTxt("verifyRand的值为" + str(verifyRand))
        loginData = {}
        loginData["account"] = "jenkins"
        loginData["password"] = "jenkinspasswordJK"
        loginData["keepLogin[]"] = "on"
        loginData["referer"] = ""
        loginData["verifyRand"] = verifyRand
        loginPage = self.session.post(url="http://192.168.10.237/zentao/user-login.html", data=loginData)
        logWriteToTxt("登录禅道")



    def get_verifyRand(self,html_str):
        """
        获取禅道登录的verifyRand_str
        :param html_str:
        :return:
        """
        pattern = re.compile(r"<input type='hidden' name='verifyRand' id='verifyRand' value='(.*)'  />")
        m = pattern.search(html_str)
        verifyRand_str = m.group(1)
        return verifyRand_str


    def findstr(self,find_str,html_str):
        """
        匹配单一字符串
        :param find_str:
        :param html_str:
        :return:
        """
        pattern = re.compile(find_str)
        m = pattern.search(html_str)
        find_str = m.group(1)
        return find_str

    def findstr2(self,findstr,html_str):
        """
        匹配多个字符串
        :param findstr:
        :param html_str:
        :return:
        """
        pattern=re.compile(findstr)
        m=pattern.search(html_str)
        list1=[]
        list1.append(m.group(1))
        list1.append(m.group(2))
        list1.append(m.group(3))
        return list1

    def get_name_list(self):
        """
        获取研发人员的名单，返回一个字典，字典的组成为'admin': ['admin', '项目经理', '男']，如下格式：
        {'admin': ['admin', '项目经理', '男'], 'yechao': ['叶超', '研发', '男'], 'gaoshihong': ['高世洪', '研发', '男'], 'duanshiyong': ['段仕勇', '高层管理', '男'], 'luyanglin': ['陆洋麟', '研发', '男'], 'qiudongsen': ['邱冬森', '项目经理', '男'], 'guoxiaoxiao': ['郭晓晓', '研发', '男'], 'luhuijun': ['陆慧君', '测试', '女'], 'yangyijun': ['杨义俊', '其他', '男'], 'lishouzhong': ['李首忠', '产品主管', '男'], 'wangzixiong': ['汪子雄', '测试', '男'], 'luojun': ['罗军', '项目经理', '男'], 'guodebin': ['郭德彬', '研发', '男'], 'zhangbin': ['张斌', '项目经理', '男'], 'zhouzipu': ['周自朴', '研发', '男'], 'suwenming': ['宿文明', '研发', '男'], 'likaiming': ['李开明', '其他', '男'], 'yanxu': ['严旭', '测试', '男'], 'qijun': ['齐俊', '高层管理', '男'], 'qining': ['齐宁', '测试', '男'], 'chenjunliang': ['陈隽樑', '测试', '男'], 'zhoutao': ['周涛', '测试', '男'], 'weiyi': ['魏毅', '其他', '男'], 'liangzonghu': ['梁宗湖', '其他', '男'], 'xuziheng': ['徐梓恒', '其他', '男'], 'yangan': ['杨桉', '其他', '男'], 'litao': ['李涛', '其他', '男'], 'ludanyong': ['陆旦勇', '其他', '男'], 'fangzihua': ['方自华', '其他', '男'], 'meihan': ['梅晗', '其他', '男'], 'fanweijian': ['樊伟健', '其他', '男'], 'geyingfeng': ['葛应峰', '研发', '男'], 'chenjinwei': ['陈进伟', '研发', '男'], 'zengqingwang': ['曾庆旺', '研发', '男'], 'liuyang': ['刘阳', '其他', '男'], 'chenweijia': ['陈维佳', '其他', '男'], 'yangmei': ['杨梅', '其他', '女'], 'chenyuan': ['陈远', '其他', '男'], 'shansirong': ['单嗣荣', '其他', '男'], 'liqingwen': ['李清文', '其他', '男'], 'suhongyu': ['苏宏宇', '其他', '男'], 'dongdaobo': ['董道波', '研发', '男'], 'liuchen': ['刘晨', '其他', '男'], 'wangyanfeng': ['王延峰', '其他', '男'], 'lijinzhong': ['李锦忠', '研发', '男'], 'qinxinchen': ['秦炘陈', '研发', '男'], 'zhanglinjun': ['张遴俊', '研发', '男'], 'dengfan': ['邓凡', '研发', '男'], 'nichengxiang': ['倪成湘', '测试', '女'], 'zengyanjun': ['曾雁军', '测试', '男'], 'dinghailei': ['丁海磊', '研发', '男'], 'panpan': ['潘攀', '研发', '男'], 'luqi': ['卢琦', '测试', '男'], 'zhanglingfei': ['张玲飞', '研发', '男'], 'zhaojiarui': ['赵嘉瑞', '研发', '男'], 'cailei': ['蔡磊', '研发', '男'], 'zhoujiezhong': ['周介忠', '研发', '男'], 'hanjushu': ['韩居舒', '测试', '女'], 'guozheng': ['郭峥', '其他', '男'], 'zhanglinghua': ['张灵华', '测试', '男'], 'zhuzhenpeng': ['朱振鹏', '测试', '男'], 'xukunlun': ['许昆仑', '其他', '男'], 'yangguibing': ['杨贵兵', '测试', '男'], 'chengxinhao': ['程鑫豪', '研发', '男'], 'renwei': ['任伟', '研发', '男'], 'liuxiaobo': ['刘晓波', '研发', '男'], 'zhushuai': ['朱帅', '研发', '男'], 'jiangtao': ['江涛', '研发', '男'], 'heyuanliang': ['何远亮', '研发', '男'], 'huangsong': ['黄宋', '其他', '男'], 'zhangchunjuan': ['张春娟', '项目经理', '女'], 'linwenbin': ['林文彬', '研发', '男'], 'qiaoxia': ['乔霞', '其他', '女'], 'hulanzhu': ['胡岚竹', '研发', '男'], 'xucaimin': ['徐采敏', '研发', '女'], 'sunyang': ['孙杨', '其他', '男'], 'wangxiaolong': ['王小龙', '研发', '男'], 'sunsihui': ['孙思慧', '研发', '男'], 'xichen': ['奚晨', '研发', '男'], 'jiangwenbo': ['蒋文波', '研发', '男'], 'caijunbo': ['蔡俊波', '其他', '男'], 'taoshaoyang': ['陶绍阳', '测试', '男'], 'xufenghua': ['徐凤华', '测试', '女'], 'wuchenchen': ['吴晨晨', '研发', '女'], 'xuxiangdong': ['徐向东', '测试', '男'], 'guoguangfei': ['郭广飞', '其他', '男'], 'yulei': ['余雷', '测试', '男'], 'wangxingdong': ['王兴东', '产品经理', '男'], 'chuyizhou': ['储一舟', '研发', '男'], 'zhanghongyuan': ['张宏远', '研发', '男'], 'baoke': ['鲍可', '研发', '男'], 'zhuconglin': ['朱从琳', '研发', '女'], 'duhao': ['杜昊', '研发', '男'], 'zhaodawei': ['赵达伟', '研发', '男'], 'dongjinxin': ['董金鑫', '研发', '男'], 'xiaopeng': ['肖鹏', '测试', '男'], 'zhaoguo': ['赵果', '研发', '男'], 'songhuan': ['宋欢', '测试', '女'], 'manlizhen': ['满丽珍', '研发', '女'], 'duanyu': ['段瑜', '其他', '女'], 'lifeng': ['李锋', '其他', '男'], 'wangguoqing': ['王国庆', '研发', '男'], 'wanghaiyang': ['王海洋', '其他', '男'], 'leixiaoju': ['雷晓菊', '测试', '女'], 'huangyuan': ['黄媛', '研发', '女'], 'zhangteng': ['张腾', '研发', '男'], 'wangshengxu': ['王胜旭', '研发', '男'], 'wanghaoran': ['汪皓然', '研发', '男'], 'liwei': ['李伟', '研发', '男'], 'yuanzheng': ['袁正', '研发', '男'], 'dongdan': ['董丹', '其他', '女'], 'lijie': ['李捷', '其他', '男'], 'cuihuailiang': ['崔怀亮', '其他', '男'], 'haoqixin': ['郝启新', '测试', '男'], 'wangbin': ['王斌', '研发主管', '男'], 'sunbin': ['孙斌', '研发', '男'], 'sunlong': ['孙龙', '研发', '男'], 'liuyanjiao': ['刘艳娇', '测试', '女'], 'wangrendong': ['王仁冬', '测试', '男'], 'xiongjinshu': ['熊锦树', '研发', '男'], 'dsy': ['段仕勇1', '产品经理', '男'], 'leichaoqun': ['雷超群', '测试', '男'], 'tangwenchao': ['唐文超', '', '男'], 'yangchunxian': ['杨春鲜', '研发', '男'], 'chuzhihao': ['初峙昊', '研发', '男'], 'wangjianhua': ['王建华', '研发', '男']}

        """
        devPage=self.session.get(url="http://192.168.10.237/zentao/company-browse-2.html")
        logWriteToTxt("进入获取研发人员名面的页面")
        dev_num=int(self.findstr(r"data-rec-total='(.*)' data-rec-per",devPage.text))
        every_page_num=int(self.findstr(r"data-rec-per-page='(.*)' data-pa",devPage.text))
        logWriteToTxt("总共有多少人"+str(dev_num))
        logWriteToTxt("每页有多少人"+str(every_page_num))
        page_num=math.ceil(dev_num/every_page_num)
        logWriteToTxt("研发人员总共有"+str(page_num)+"页")
        dev_name_dict={}
        for i in range(1,page_num+1):
            content=self.session.get(url="http://192.168.10.237/zentao/company-browse-2-bydept-id-"+str(dev_num)+"-20-"+str(i)+".html")
            logWriteToTxt("对第"+str(i)+"页进行解析")
            soup = BeautifulSoup(content.text, 'lxml')
            userlist=soup.select("tr")
            for index,tr in enumerate(userlist):
                if index!=0:
                    html_str=str(tr)
                    name=self.findstr(r'type="checkbox" value="(.*?)"/>',html_str)
                    chinese_name=self.findstr(r'html" title="(.*?)">',html_str)
                    title=self.findstr(r'class="w-90px" title="(.*?)">',html_str)
                    gender=self.findstr(r'<td class="c-type">(.*?)</td>',html_str)
                    # print(name)
                    # print(chinese_name)
                    # print("********")
                    dev_info=[]
                    dev_info.append(chinese_name)
                    dev_info.append(title)
                    dev_info.append(gender)
                    dev_name_dict[name]=dev_info
        return dev_name_dict

    def get_department(self):
        """
        获取各个部门的人员信息，返回一个字典列表，如下格式：
        {'总裁办': {'guodebin': ['郭德彬', '研发', '男'], 'qijun': ['齐俊', '高层管理', '男'], 'yangmei': ['杨梅', '其他', '女']}, '平台部': {'admin': ['admin', '项目经理', '男'], 'gaoshihong': ['高世洪', '研发', '男'], 'luyanglin': ['陆洋麟', '研发', '男'], 'dongdaobo': ['董道波', '研发', '男'], 'zhanglinjun': ['张遴俊', '研发', '男'], 'panpan': ['潘攀', '研发', '男'], 'zhangchunjuan': ['张春娟', '项目经理', '女'], 'linwenbin': ['林文彬', '研发', '男'], 'xucaimin': ['徐采敏', '研发', '女'], 'wuchenchen': ['吴晨晨', '研发', '女'], 'zhaoguo': ['赵果', '研发', '男'], 'liwei': ['李伟', '研发', '男']}, '热点产品部': {'yechao': ['叶超', '研发', '男'], 'lishouzhong': ['李首忠', '产品主管', '男'], 'suwenming': ['宿文明', '研发', '男'], 'yanxu': ['严旭', '测试', '男'], 'geyingfeng': ['葛应峰', '研发', '男'], 'chenjinwei': ['陈进伟', '研发', '男'], 'zengqingwang': ['曾庆旺', '研发', '男'], 'dengfan': ['邓凡', '研发', '男'], 'nichengxiang': ['倪成湘', '测试', '女'], 'zengyanjun': ['曾雁军', '测试', '男'], 'zhaojiarui': ['赵嘉瑞', '研发', '男'], 'cailei': ['蔡磊', '研发', '男'], 'zhoujiezhong': ['周介忠', '研发', '男'], 'hanjushu': ['韩居舒', '测试', '女'], 'zhanglinghua': ['张灵华', '测试', '男'], 'yangguibing': ['杨贵兵', '测试', '男'], 'chengxinhao': ['程鑫豪', '研发', '男'], 'renwei': ['任伟', '研发', '男'], 'liuxiaobo': ['刘晓波', '研发', '男'], 'hulanzhu': ['胡岚竹', '研发', '男']}, '反制产品部': {'guoxiaoxiao': ['郭晓晓', '研发', '男'], 'luojun': ['罗军', '项目经理', '男'], 'qinxinchen': ['秦炘陈', '研发', '男'], 'zhuzhenpeng': ['朱振鹏', '测试', '男'], 'taoshaoyang': ['陶绍阳', '测试', '男'], 'xufenghua': ['徐凤华', '测试', '女'], 'leixiaoju': ['雷晓菊', '测试', '女'], 'leichaoqun': ['雷超群', '测试', '男'], 'tangwenchao': ['唐文超', '', '男']}, '路测产品部': {'qiudongsen': ['邱冬森', '项目经理', '男'], 'luqi': ['卢琦', '测试', '男'], 'heyuanliang': ['何远亮', '研发', '男'], 'zhaodawei': ['赵达伟', '研发', '男'], 'wangshengxu': ['王胜旭', '研发', '男'], 'yuanzheng': ['袁正', '研发', '男']}, '大数据产品部': {'duanshiyong': ['段仕勇', '高层管理', '男'], 'luhuijun': ['陆慧君', '测试', '女'], 'zhangbin': ['张斌', '项目经理', '男'], 'zhouzipu': ['周自朴', '研发', '男'], 'qining': ['齐宁', '测试', '男'], 'lijinzhong': ['李锦忠', '研发', '男'], 'zhanglingfei': ['张玲飞', '研发', '男'], 'zhushuai': ['朱帅', '研发', '男'], 'jiangtao': ['江涛', '研发', '男'], 'jiangwenbo': ['蒋文波', '研发', '男'], 'yulei': ['余雷', '测试', '男'], 'chuyizhou': ['储一舟', '研发', '男'], 'baoke': ['鲍可', '研发', '男'], 'duhao': ['杜昊', '研发', '男'], 'dongjinxin': ['董金鑫', '研发', '男'], 'huangyuan': ['黄媛', '研发', '女'], 'zhangteng': ['张腾', '研发', '男'], 'wanghaoran': ['汪皓然', '研发', '男'], 'liuyanjiao': ['刘艳娇', '测试', '女'], 'wangrendong': ['王仁冬', '测试', '男']}, '可视化产品部': {'dinghailei': ['丁海磊', '研发', '男'], 'wangxiaolong': ['王小龙', '研发', '男'], 'sunsihui': ['孙思慧', '研发', '男'], 'xichen': ['奚晨', '研发', '男'], 'wangxingdong': ['王兴东', '产品经理', '男'], 'zhuconglin': ['朱从琳', '研发', '女'], 'songhuan': ['宋欢', '测试', '女'], 'haoqixin': ['郝启新', '测试', '男'], 'chuzhihao': ['初峙昊', '研发', '男']}, '合肥系统部': {'wangbin': ['王斌', '研发主管', '男'], 'sunbin': ['孙斌', '研发', '男'], 'sunlong': ['孙龙', '研发', '男'], 'xiongjinshu': ['熊锦树', '研发', '男']}, '运营部': {'suhongyu': ['苏宏宇', '其他', '男'], 'qiaoxia': ['乔霞', '其他', '女'], 'duanyu': ['段瑜', '其他', '女']}, '市场部': {'yangyijun': ['杨义俊', '其他', '男'], 'wangzixiong': ['汪子雄', '测试', '男'], 'chenweijia': ['陈维佳', '其他', '男'], 'chenyuan': ['陈远', '其他', '男'], 'wangyanfeng': ['王延峰', '其他', '男'], 'xukunlun': ['许昆仑', '其他', '男'], 'dongdan': ['董丹', '其他', '女'], 'lijie': ['李捷', '其他', '男']}, '销售部': {'meihan': ['梅晗', '其他', '男'], 'liuyang': ['刘阳', '其他', '男']}, '售后服务部': {'likaiming': ['李开明', '其他', '男'], 'zhoutao': ['周涛', '测试', '男'], 'weiyi': ['魏毅', '其他', '男'], 'liangzonghu': ['梁宗湖', '其他', '男'], 'xuziheng': ['徐梓恒', '其他', '男'], 'yangan': ['杨桉', '其他', '男'], 'litao': ['李涛', '其他', '男'], 'ludanyong': ['陆旦勇', '其他', '男'], 'fangzihua': ['方自华', '其他', '男'], 'fanweijian': ['樊伟健', '其他', '男'], 'liuchen': ['刘晨', '其他', '男'], 'guozheng': ['郭峥', '其他', '男'], 'huangsong': ['黄宋', '其他', '男'], 'sunyang': ['孙杨', '其他', '男'], 'caijunbo': ['蔡俊波', '其他', '男'], 'guoguangfei': ['郭广飞', '其他', '男'], 'lifeng': ['李锋', '其他', '男'], 'wanghaiyang': ['王海洋', '其他', '男'], 'cuihuailiang': ['崔怀亮', '其他', '男']}, '生产部': {'chenjunliang': ['陈隽樑', '测试', '男'], 'shansirong': ['单嗣荣', '其他', '男'], 'liqingwen': ['李清文', '其他', '男']}}

        :return:
        """
        content=self.session.get(url="http://192.168.10.237/zentao/company-browse-2.html")
        soup=BeautifulSoup(content.text,"html.parser")
        department=soup.select("#sidebar > div.cell > ul ")
        list1=str(department).split("\n")
        find_str=r'<li><a href="(.*)" id="(.*)">(.*)</a>'
        dep_dict={}

        for str1 in list1:
            try:
                dep_info=self.findstr2(find_str,str1)
                if dep_info[2] != '研发体系':
                    devPage = self.session.get(url="http://192.168.10.237" + dep_info[0])
                    logWriteToTxt("对" + str(dep_info[2]) + "部门进行解析")
                    dev_num = int(self.findstr(r"data-rec-total='(.*)' data-rec-per", devPage.text))
                    every_page_num = int(self.findstr(r"data-rec-per-page='(.*)' data-pa", devPage.text))
                    logWriteToTxt("总共有多少人" + str(dev_num))
                    logWriteToTxt("每页有多少人" + str(every_page_num))
                    page_num = math.ceil(dev_num / every_page_num)
                    logWriteToTxt("研发人员总共有" + str(page_num) + "页")
                    dev_name_dict = {}
                    url_list=dep_info[0].split(".")
                    for i in range(1, page_num + 1):
                        content = self.session.get(
                            url="http://192.168.10.237"+url_list[0]+"-bydept-id-"+str(dev_num)+"-20-" + str(i) + ".html")
                        logWriteToTxt("对第" + str(i) + "页进行解析")
                        soup = BeautifulSoup(content.text, 'lxml')
                        userlist = soup.select("tr")
                        for index, tr in enumerate(userlist):
                            if index != 0:
                                html_str = str(tr)
                                name = self.findstr(r'type="checkbox" value="(.*?)"/>', html_str)
                                chinese_name = self.findstr(r'html" title="(.*?)">', html_str)
                                title = self.findstr(r'class="w-90px" title="(.*?)">', html_str)
                                gender = self.findstr(r'<td class="c-type">(.*?)</td>', html_str)
                                dev_info = []
                                dev_info.append(chinese_name)
                                dev_info.append(title)
                                dev_info.append(gender)
                                dev_name_dict[name] = dev_info
                    dep_dict[dep_info[2]] = dev_name_dict
            except:
                pass
        return dep_dict


if __name__=="__main__":
    a=Zendao()
    # print(a.get_name_list())
    # print(a.get_department())
    # i=0
    # print(a.get_department())
    # for key in a.get_department().keys():
    #     print(key)
    #     i=i+1
    # print(i)

    print(list(a.get_name_list().keys()))


    # department_info = a.get_department()
    # department_list = department_info.keys()
    # name_department_list = []
    # for department_key in department_list:
    #     if department_key in ['平台部', '热点产品部', '反制产品部', '路测产品部', '大数据产品部', '可视化产品部', '安检产品部']:
    #         demo_department_info = department_info.get(department_key)
    #         for name_item in demo_department_info.keys():
    #             if name_item != 'admin':
    #                 name_info_list = []
    #                 name_info_list.append(name_item)
    #                 name_info_list.append(demo_department_info.get(name_item)[0])
    #                 name_info_list.append(department_key)
    #                 name_department_list.append(name_info_list)
    # print(name_department_list)
    # print(len(name_department_list))












