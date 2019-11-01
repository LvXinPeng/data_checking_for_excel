import datetime
import sys
import time

import pandas as pd
import re
#
# def to_upper(row_idx, col, val):
#     if str(df_fcst[header_std[col]][row_idx]).upper() == val.upper():
#         df_fcst[header_std[col]][row_idx] = val
#

# match = re.search(pattern="^(?:[1-9][0-9]*(?:\.[0-9]+)?|0(?:\.[0-9]+)?)$", string='')
# print(match)
import tqdm
from numpy import nan

start = time.time()
def correct_number(x):
    x = str(x).replace(' ', '')
    if x:
        x = float(re.search('\d+(\.\d+)?', str(x)).group())
    else:
        x = 0
    return x

#
# def calc_total_budget(row_idx):
#     total_budget = 0.0
#     for col_idx in range(11):
#         if df_fcst[header_std[col_idx + 20]][row_idx] == ' ':
#             total_budget = total_budget * -1
#             # print(header_std[col_idx + 20])
#             print(df_fcst[header_std[col_idx + 20]][row_idx])
#         else:
#             total_budget = total_budget + df_fcst[header_std[col_idx + 20]][row_idx]
#             # print(df_fcst[header_std[col_idx + 20]][row_idx])
#
#             # print(header_std[col_idx + 20])
#             # print('++++++++++++++++')
#     return total_budget


header_std = ['Issue', 'Dept', 'Team', 'Carline', 'Lifecycle', 'Branding/NonBranding',
              'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type',
              'Activity', 'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order',
              'KPI Others', 'SMM Campaign Code (Y/N)', 'SC NO.', ' SC Name',
              'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
              'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']

#
# print(str(list(df_fcst.columns[0:34])).upper() == str(header_std).upper())
# print(str(list(df_fcst.columns[0:34])).upper())
# print(str(header_std).upper())
#
# a = []
# b = [1,2]
# c = 3
# a.append(b)
# a.append(b)
# print(a)
# a.append(c)
# print(a)
# file_log = "D:/asp_up.xlsx"
# print(datetime.datetime.now())
# log_time = []
# log_col = ["LastUpdatedTime"]
# log_time.append(str(datetime.datetime.now()))
# print(log_time)
# writer_log = pd.ExcelWriter(path=file_log, mode='w', engine='xlsxwriter')
# frame_log = pd.DataFrame(data=log_time, columns=log_col)
# frame_log.to_excel(writer_log, index=False, header=True)
# writer_log.save()
# writer_log.close()
# idx_row_data = df_fcst[header_std[1]][0]
# print("col 2:[" + str(idx_row_data) + "]")
# flag = 0
# for idx_row in range(len(df_fcst[header_std[0]])):
#     if calc_total_budget(idx_row) != df_fcst[header_std[32]][idx_row]:
#         flag = flag + 1
#         # print('--------')
#         # print(flag)
#         # print(idx_row)
#         # print('**************')
#
# # print(flag)
# #
# # df_fcst['Jan'] = df_fcst['Jan'].apply(lambda x: correct_number(x))
# for col_idx in range(12):
#     print(col_idx)
#     print(header_std[col_idx + 20])
#
# print(header_std[12 + 20])
#
#
# df_fcst = pd.read_excel('D:/asp.xlsx')
# file_log = "D:/ASP Database & actual 0906 - clear_modified.xlsx"
#
# for idx_row in range(len(df_fcst[header_std[0]])):
#     to_upper(idx_row, 5, 'Branding')
#     to_upper(idx_row, 5, 'NonBranding')
#     to_upper(idx_row, 6, 'Working')
#     to_upper(idx_row, 6, 'NonWorking')
#
# writer_log = pd.ExcelWriter(path=file_log, mode='w', engine='xlsxwriter')
# frame_log = pd.DataFrame(data=df_fcst, columns=header_std)
# frame_log.to_excel(writer_log, index=False, header=True,sheet_name='7+5 Detail')
# writer_log.save()
# writer_log.close()

header_std = ['Issue', 'Dept', 'Team', 'Carline', 'Lifecycle', 'Branding/NonBranding',
              'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type',
              'Activity', 'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order',
              'KPI Others', 'SMM Campaign Code (Y/N)', 'SC NO.', ' SC Name',
              'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
              'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
# 异常数据excel header
header_err = ['File Name', 'Exception Type', 'Index', 'Issue', 'Dept', 'Team', 'Carline', 'Lifecycle',
              'Branding/NonBranding', 'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type', 'Activity',
              'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order', 'KPI Others', 'SMM Campaign Code (Y/N)',
              'SC No.', ' SC Name', 'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct',
              'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']

# print(len(header_std))
# print(len(header_err))
#
# print('File Name'.replace(' ','').upper() in str(list(['File Name', 'Exception Type', 'Index', 'Issue'])).replace(' ','').upper())

today = str(datetime.date.today()).replace('-', '')
# print("MEDIATEAM" in ["ABBMEDIATEAM"])
# print(str(list(['Fleet Team', 'Used Car Team', 'Financial Service Team'])).replace(' ', '').upper())
# print('Fleet Team'.replace(' ', '').upper() in str(list(['Fleet', 'Used Car Team', 'Financial Service Team'])).replace(' ', '').upper())
# print('media team'.replace(' ', '').upper() in ['TRADITIONALMEDIATEAM','DIDMEDIATEAM'])
path = "D:/ASP - Erin/Raw Data/"
file_path_a = "D:/ASP - Erin/FCST/Data&Log/"
file_path_b = "D:/ASP - Erin/Header Error/"
file_path_c = "D:/ASP - Erin/Header Error/Files/"
file_path_d = "D:/ASP - Erin/Actual/Data&Log/"

file_log = "D:/ASP - Erin/Log.xlsx"
file_new = "D:/ASP - Erin/Actual/Data&Log/Act_Summary_" + today + ".xlsx"
file_new_fcst = "D:/ASP - Erin/FCST/Data&Log/FCST_Summary_" + today + ".xlsx"
file_err = "D:/ASP - Erin/Actual/Data&Log/Act_Error_" + today + ".xlsx"
file_err_fcst = "D:/ASP - Erin/FCST/Data&Log/FCST_Error_" + today + ".xlsx"
file_header_err = "D:/ASP - Erin/Header Error/Header Error_" + today + ".xlsx"

path_fcst = "D:/ASP - Erin/FCST/Dept-Team/"
file_summary_fcst = "D:/ASP - Erin/FCST/Data&Log/FCST_Summary_" + today + ".xlsx"
file_MKT = "D:/ASP - Erin/FCST/Dept-Team/FCST_MKT_" + today + ".xlsx"
file_MKT_Launch = "D:/ASP - Erin/FCST/Dept-Team/FCST_MKT Launch_" + today + ".xlsx"
file_APAC_CC = "D:/ASP - Erin/FCST/Dept-Team/FCST_APAC CC_" + today + ".xlsx"
file_Customer_Service = "D:/ASP - Erin/FCST/Dept-Team/FCST_Customer Service_" + today + ".xlsx"
file_DTS = "D:/ASP - Erin/FCST/Dept-Team/FCST_DTS_" + today + ".xlsx"
file_MI = "D:/ASP - Erin/FCST/Dept-Team/FCST_MI_" + today + ".xlsx"
file_Sales_MKT1 = "D:/ASP - Erin/FCST/Dept-Team/FCST_Sales MKT_DMKT_" + today + ".xlsx"
file_Sales_MKT2 = "D:/ASP - Erin/FCST/Dept-Team/FCST_Sales MKT_HK_" + today + ".xlsx"
file_Sales_MKT3 = "D:/ASP - Erin/FCST/Dept-Team/FCST_Sales MKT_Internal Fleet Car_" + today + ".xlsx"
file_Sales_MKT4 = "D:/ASP - Erin/FCST/Dept-Team/FCST_Sales MKT_National Sales_" + today + ".xlsx"
file_Sales_MKT5 = "D:/ASP - Erin/FCST/Dept-Team/FCST_Sales MKT_NBD Team_" + today + ".xlsx"



path_actual = "D:/ASP - Erin/Actual/Dept-Team/"
file_summary_actual = "D:/ASP - Erin/Actual/Data&Log/Act_Summary_" + today + ".xlsx"
file_MKT2 = "D:/ASP - Erin/Actual/Dept-Team/Act_MKT_" + today + ".xlsx"
file_MKT_Launch2 = "D:/ASP - Erin/Actual/Dept-Team/Act_MKT Launch_" + today + ".xlsx"
file_APAC_CC2 = "D:/ASP - Erin/Actual/Dept-Team/Act_APAC CC_" + today + ".xlsx"
file_Customer_Service2 = "D:/ASP - Erin/Actual/Dept-Team/Act_Customer Service_" + today + ".xlsx"
file_DTS2 = "D:/ASP - Erin/Actual/Dept-Team/Act_DTS_" + today + ".xlsx"
file_MI2 = "D:/ASP - Erin/Actual/Dept-Team/Act_MI_" + today + ".xlsx"
file_Sales_MKT12 = "D:/ASP - Erin/Actual/Dept-Team/Act_Sales MKT_DMKT_" + today + ".xlsx"
file_Sales_MKT22 = "D:/ASP - Erin/Actual/Dept-Team/Act_Sales MKT_HK_" + today + ".xlsx"
file_Sales_MKT32 = "D:/ASP - Erin/Actual/Dept-Team/Act_Sales MKT_Internal Fleet Car_" + today + ".xlsx"
file_Sales_MKT42 = "D:/ASP - Erin/Actual/Dept-Team/Act_Sales MKT_National Sales_" + today + ".xlsx"
file_Sales_MKT52 = "D:/ASP - Erin/Actual/Dept-Team/Act_Sales MKT_NBD Team_" + today + ".xlsx"
file_std = "D:/ASP - Erin/Standard.xlsx"
standard = pd.read_excel(file_std, sheet_name='Folder Path', dtype=str)
dept_std = pd.read_excel(file_std, sheet_name='Department Standard', dtype=str)
carline_std = pd.read_excel(file_std, sheet_name='Carline Standard', dtype=str)
lifecycle_std = pd.read_excel(file_std, sheet_name='Lifecycle Standard', dtype=str)
nonworking_std = pd.read_excel(file_std, sheet_name='NonWorking Standard', dtype=str)
working_std = pd.read_excel(file_std, sheet_name='Working Standard', dtype=str)
# dept_team_std = pd.read_excel(file_std, sheet_name='Dept & Team', dtype=str)

def read_dept(std):
    result = []
    for item in list(std.columns):
        if str(item) != 'nan':
            result.append(str(item).replace(' ', '').upper())
    return result


def read_std(std):
    result = []
    for item in list(std[std.columns[0]]):
        if str(item) != 'nan':
            result.append(str(item).replace(' ', '').upper())
    return result


def read_sub_std(std, col):
    result = []
    col_upper = str(col)[0].upper() + str(col)[1:]
    for item in list(std[col_upper.strip()]):
        if str(item) != 'nan':
            result.append(str(item).replace(' ', '').upper())
    return result


def read_sub_std_plus(std, col):
    result = []
    col_upper = str(col)[0].upper() + str(col)[1:]
    for item in list(std[col_upper.strip()]):
        if str(item) != 'nan':
            result.append(str(item).replace(' ', '').upper())
    return result

df_folder_path = pd.read_excel(file_std, sheet_name='Folder Path', dtype=str)
folder_path = df_folder_path['Path'][0]
df_issue = pd.read_excel(file_std, sheet_name='Issue Standard', dtype=str)
issue_list = []
for item in df_issue['Issue FCST']:
    if str(item) != 'nan':
        issue_list.append(item.replace(' ', '').upper())
# print(issue_list)
# print(issue_list.remove('ACTUAL'))
hk = ['SALESMKT', 'HK']
# print(read_std(standard))

def progress(timeout=10):
    timeout = round(timeout)
    if timeout < 1:
        timeout = 1
    for i in range(timeout):
        pro = round((i + 1) / timeout * 100.0)
        sys.stdout.write("\r%s%s[%d%%] " % ("█" * pro, " " * (100 - pro), pro))
        sys.stdout.flush()
        time.sleep(1)
    # break line
    print()
start = time.time()

# progress(int(time.time() - start))

# for i in tqdm.trange(100):
#     time.sleep(0.1)
#
# # print('hello')
# sys.stdout.write("%s%d%s" % ("It has taken ", int(time.time() - start)," seconds"))
# sys.stdout.flush()
# print(read_sub_std(dept_team_std, 'Sales MKT'))
# print(read_sub_std(dept_team_std, 'Sales MKT'))
# print(hk[0] in read_dept(dept_team_std))
# print(hk[1] in read_sub_std(dept_team_std, 'Sales MKT'))
# print(hk[0] in read_dept(dept_team_std) and hk[1] in read_sub_std(dept_team_std, 'Sales MKT'))

#
# print(read_std(file_std, 'Carline Standard'))
# print(read_dept(file_std, 'Department Standard'))
# for item in read_dept(file_std, 'Department Standard'):
#     print(read_dept_std(file_std, 'Department Standard', item))
print('--------------------------------')
#
# standard_department = pd.read_excel(file_std, sheet_name='Department Standard', dtype=str)
# print(list(standard_department.columns))
# for i in list(standard_department.columns):
#     print(i)
#     print(list(standard_department[i]))
# print('--------------------------------')
# standard_carline = pd.read_excel(file_std, sheet_name='Carline Standard', dtype=str)
# print(list(standard_carline[standard_carline.columns[0]]))
# print('--------------------------------')
#
# standard_lifecycle = pd.read_excel(file_std, sheet_name='Lifecycle Standard', dtype=str)
# print(list(standard_lifecycle[standard_lifecycle.columns[0]]))
# print('--------------------------------')
# standard_working = pd.read_excel(file_std, sheet_name='Working Standard', dtype=str)
# print(list(standard_working['Sales funnel']))
# print(list(standard_working['Category']))
# try:
#     for i in list(standard_working['Category']):
#         print(i)
#         print(list(standard_working[i]))
# except KeyError:
#     pass
# print('--------------------------------')
# standard_nonworking = pd.read_excel(file_std, sheet_name='NonWorking Standard', dtype=str)
# print(list(standard_nonworking['Sales funnel']))
# print(list(standard_nonworking['Category']))
# try:
#     for i in list(standard_nonworking['Category']):
#         print(i)
#         print(list(standard_nonworking[i]))
# except KeyError:
#     pass
cost = time.time() - start
print(str(int(cost)) + "s")
old_string = 'Sales MKT'
pattern="[A-Z]"
if old_string.strip().find(' ') == -1:
    new_string=re.sub(pattern,lambda x:" "+x.group(0),old_string)
    print(new_string.strip())
alpha = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
print(list(old_string))
print(old_string not in alpha)
print(old_string.isupper())
new_str = ''
for i in range(len(old_string)):
    if 65 <= ord(old_string[i]) <= 90:
        if i != 0 and 97 <= ord(old_string[i-1]) <= 122:
            new_str = new_str + ' '
        new_str = new_str + old_string[i]
    else:
        new_str = new_str + old_string[i]
print(new_str)