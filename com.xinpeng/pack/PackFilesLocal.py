import datetime
import os
import zipfile
import pandas as pd
import shutil


def zipDir(dirpath, outFullName):
    # 压缩指定文件夹
    # :param dirpath: 目标文件夹路径
    # :param outFullName: 压缩文件保存路径+xxxx.zip
    # :return: 无
    zip = zipfile.ZipFile(outFullName, "w", zipfile.ZIP_DEFLATED)
    for path, dirnames, filenames in os.walk(dirpath):
        # 去掉目标跟路径，只对目标文件夹下边的文件及文件夹进行压缩
        fpath = path.replace(dirpath, '')
        for filename in filenames:
            zip.write(os.path.join(path, filename), os.path.join(fpath, filename))
    zip.close()


def rmdir(srcDir):
    file_list = os.listdir(srcDir)  # 列出该目录下的所有文件名
    for f in file_list:
        file_path = os.path.join(srcDir, f)  # 将文件名映射成绝对路劲
        if os.path.isfile(file_path):  # 判断该文件是否为文件或者文件夹
            os.remove(file_path)  # 若为文件，则直接删除
        elif os.path.isdir(file_path):
            shutil.rmtree(file_path, True)  # 若为文件夹，则删除该文件夹及文件夹内所有文件


def mkdir(file_path):
    if not os.path.exists(file_path):
        os.makedirs(file_path)


year = str(datetime.date.today().year)
today = str(datetime.date.today())
file_log = "D:/ASP - Erin/Log.xlsx"
srcPath = 'D:/ASP - Erin/'
try:
    issue_val = pd.read_excel(file_log, dtype=str)['Issue'][0]
    destPath = 'D:/Package/' + issue_val + ' ' + year + '.zip'
except:
    destPath = 'D:/Package/'+ today + '.zip'
zipDir(srcPath, destPath)
rmdir(srcPath)
mkdir(srcPath + 'Raw data')
print("Process finished with exit code 0")