# -*- coding: utf-8 -*- 

"""
# @Time : 2019/3/28
# @Author : xucaimin
"""
import os
import sys
import shutil

#获取文件件下所有的文件名称、目录等
def getfileName(file_dir):
    for root, dirs, files in os.walk(file_dir):
        # print(1,root)  # 当前目录路径
        # print(2,dirs)  # 当前路径下所有子目录
        print(3,files)  # 当前路径下所有非目录子文件
    # for index,file in enumerate(files):
    #     print(index,path+"/"+file)
    # print(len(files))
    return files


#获取文件夹下的最新文件夹或者文件
def new_file(test_file):
    lists = os.listdir(test_file)         # 列出目录的下所有文件和文件夹保存到lists
    lists.sort(key=lambda fn: os.path.getmtime(test_file + "/" + fn)) # 按时间排序
    file_new = os.path.join(test_file, lists[-1])      # 获取最新的文件保存到file_new
    # print("1,",file_new)
    return file_new

def get_last_file(test_file,basename):
    print(basename)
    lists = os.listdir(test_file)  # 列出目录的下所有文件和文件夹保存到lists
    # print(lists)
    file_path=''
    for a in lists:
        # print(str(a))
        if basename.strip() in str(a).strip() :
            file_path=os.path.join(test_file,a)
            print("比对的文件夹"+file_path)
    return file_path

    # print("1,",file_new)
    # return file_new
#得到所有文件夹下的特定文件
def get_str_new_file(test_file,FlagStr):
    FileList = []
    FileNames = os.listdir(test_file)
    # print(FileNames)
    if (len(FileNames) > 0):
        for fn in FileNames:
            if (len(FlagStr) > 0):
                if(FlagStr in fn):
                    FileList.append(fn)
    if(len(FileList)!=0):
        FileList.sort(key=lambda fn: os.path.getmtime(test_file + "/" + fn))  # 按时间排序
        file_new = os.path.join(test_file, FileList[-1])  # 获取最新的文件保存到file_new
        # print("1,",file_new)
        # print(FileList[-1])
        # return FileList[-1]
        return file_new

#得到当前项目的目录
def getPath():
    path=sys.path[0]#如果PYTHONPATH 变量还不存在，可以创建它！路径会自动加入到sys.path中
    return path


#删除文件
def del_files(path):
    if os.path.exists(path):  # 如果文件存在
        # 删除文件，可使用以下两种方法。
        os.remove(path)  # 则删除
        # os.unlink(my_file)
    else:
        print('no such file:%s' % path)


#将文件夹下的所有文件都复制到另一文件夹
def copy_search_file(srcDir, desDir):
    ls = os.listdir(srcDir)
    for line in ls:
        filePath = os.path.join(srcDir, line)
        if os.path.isfile(filePath):
            print (filePath)
            shutil.copy(filePath, desDir)

#将文件夹下的特定文件都复制到另一文件夹,并且把原先的文件删掉，达到剪切的效果
def copy__file(filePath, desDir):
    if os.path.isfile(filePath):
        print("下载文件中最新的是:",filePath)
        shutil.copy(filePath, desDir)
        os.remove(filePath)


if __name__=="__main__":
    computerName = os.getlogin()
    downloadPath = "C:\\Users\\" + computerName + "\\Downloads"
    # path=os.path.dirname(os.getcwd())+ '/result'
    # print(path)
    # filepath=new_file(downloadPath)
    # print(filepath)
    # getfileName(filepath)
    print(get_str_new_file(downloadPath,".csv"))

