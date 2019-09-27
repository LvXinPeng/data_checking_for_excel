import datetime

import pandas as pd
import re
#
# def to_upper(row_idx, col, val):
#     if str(df_fcst[header_std[col]][row_idx]).upper() == val.upper():
#         df_fcst[header_std[col]][row_idx] = val
#

# match = re.search(pattern="^(?:[1-9][0-9]*(?:\.[0-9]+)?|0(?:\.[0-9]+)?)$", string='')
# print(match)


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

print(len(header_std))
print(len(header_err))

print('File Name'.replace(' ','').upper() in str(list(['File Name', 'Exception Type', 'Index', 'Issue'])).replace(' ','').upper())

today = str(datetime.date.today()).replace('-', '')

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