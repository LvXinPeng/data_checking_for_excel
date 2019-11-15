# for idx in range(len(agent_code_unique)):
# print(str(agent_code_unique[idx]) + "-三级电气技师-" + str(
#     df[(df['三级电气技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()))
# print(str(agent_code_unique[idx]) + "-三级机械技师-" + str(
#     df[(df['三级机械技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()))
# print(str(agent_code_unique[idx]) + "-三级空调技师-" + str(
#     df[(df['三级空调技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()))
# print(str(agent_code_unique[idx]) + "-四级技师-" + str(
#     df[(df['四级技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()))
# print("==============")
# g3elec = df[(df['三级电气技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
# g3mech = df[(df['三级机械技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
# g3airc = df[(df['三级空调技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
# g4tech = df[(df['四级技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
# print(g3elec)
# print(g3mech)
# print(g3airc)
# print(g4tech)
# print("***********************")

#
# print('V700 – V900'.split(' – V')[0].split('V')[1])
# # print(len( 'D:/ASP - Erin/Raw Data/ASP Dbase Report 2019 FCST 1+11.xlsx'.split('.')[0][-9:]))
import os
#
# print('            -'.replace('-', '').replace(' ','') == '')
# listdir =  os.listdir('D:/ASP - Erin/Raw Data/')
# file = 'D:/ASP - Erin/Raw Data/' + listdir[0]
# print(file)
# print(listdir)
import stat

from threading import Timer
import  time
#
# def show():
#     cost = time.time() - 0
#     ct = str(cost) + "s"
#     print(ct)
#
#
# # 指定一秒钟之后执行 show 函数
# t = Timer(0.5, show)
# t.start()
# res = 0
# for i in range(1000000):
#     res = i + res

import pandas as pd

import threading, time

global t
start = time.time()

# def show_time():
#     print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
#     t = threading.Timer(1.0, show_time)
#     t.start()
#
import datetime

# show_time()
# # t=threading.Timer(1.0,sayHello)
# # t.start()
#
# for i in range(100):
#
# # os.system("exit")
#     print("The program is running and it has been running for %.2fseconds"%(time.time()-start))
#     os.system('cls')
# print(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time())))
cur_year = datetime.datetime.now().year
cur_month = datetime.datetime.now().month
month_list = []
for idx_mon in range(6):
    month = cur_month + idx_mon + 1
    month_temp = ''
    if month <= 9:
        month_temp = str(cur_year) + '0' + str(month)
    elif 10 <= month <= 12:
        month_temp = str(cur_year) + str(month)
    elif month > 12:
        month_temp = str(cur_year + 1) + '0' + str(abs(month - 12))
    month_list.append(month_temp)
# for idx_month in month_list:
#     file_path = "D:/" + str(idx_month)
#     if not os.path.exists(file_path):
#         os.makedirs(file_path)

print(cur_year)
print(cur_month)
print(month_list)

file_std = 'D:/Projects/Volvo BI Project/ASP - Erin/Standard.xlsx'
df_folder_path = pd.read_excel(file_std, sheet_name='Folder Path', dtype=str)
# print(os.listdir(df_folder_path['Path'][0]))
path = df_folder_path['Path'][0]
# print(path)
    #'D:/Projects/Volvo BI Project/ASP - Erin/201911/'
def change_dir_before_excute(src_path):
    path_list = ['Result/Checking/Actual/Data&Log/', 'Result/Checking/FCST/Data&Log/', 'Result/Checking/Budget/Data&Log/']
    for path_item in path_list:
        temp_path = src_path + path_item
        try:
            os.listdir(temp_path)
        except FileNotFoundError:
            temp_path = src_path + '/' + path_item
            # print(temp_path)
        if os.listdir(temp_path):
            for item in os.listdir(temp_path):
                print(item)
                os.chmod(temp_path + item, stat.S_IWRITE)

change_dir_before_excute(path)
try:
    print(os.listdir(path + 'Result/Checking/FCST/Data&Log/'))
except FileNotFoundError:
    path = path + '/'

print(os.listdir(path + 'Result/Checking/FCST/Data&Log/'))