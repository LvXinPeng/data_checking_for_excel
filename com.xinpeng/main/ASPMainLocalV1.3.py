import time
import pandas as pd
import os
import shutil
import stat
import sys


def mkdir(file_path):
    if not os.path.exists(file_path):
        os.makedirs(file_path)


def is_null(df, header, row_idx):
    is_na = True
    if str(df[header[0]][row_idx]) == "nan" or str(df[header[1]][row_idx]) == "nan":
        is_na = False
    return is_na


def is_ssm(df, header, row_idx):
    is_ssmc = True
    if str(df[header[16]][row_idx]).strip() not in ['Y', 'y', 'N', 'n', 'N.A.']:
        is_ssmc = False
    else:
        if str(df[header[16]][row_idx]).strip() in ['N.A.'] \
                and str(df[header[17]][row_idx]) not in ['N.A.', '0']:
            is_ssmc = False
        elif str(df[header[16]][row_idx]).strip() in ['Y', 'y'] \
                and str(df[header[17]][row_idx]) in ['N.A.', '0']:
            is_ssmc = False
    return is_ssmc


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
    col_upper = str(col)[0].upper() + str(col).strip()[1:]
    for item in list(std[col_upper]):
        if str(item) != 'nan':
            result.append(str(item).replace(' ', '').upper())
    return result


def is_life_cycle(df, header, row_idx):
    is_life = False
    std = read_std(lifecycle_std)
    if str(df[header[4]][row_idx]).replace(' ', '').upper() in std:
        is_life = True
    return is_life


def is_carline(df, header, row_idx):
    is_car = False
    std = read_std(carline_std)
    if str(df[header[3]][row_idx]).replace(' ', '').upper() in std:
        is_car = True
    return is_car


def is_branding(df, header, row_idx):
    is_b = False
    if str(df[header[5]][row_idx]).replace(' ', '').upper() in ['BRANDING', 'NONBRANDING']:
        is_b = True
    return is_b


def calc_total_budget(df, header, row_idx):
    temp = []
    for col_idx in range(12):
        try:
            temp.append(float(df[header[col_idx + 20]][row_idx]))
        except ValueError:
            temp.append(0.0)
    total_budget = sum(list(map(float, temp)))
    is_eq = total_budget == float(df[header[32]][row_idx]) or abs(total_budget - float(df[header[32]][row_idx])) <= 1
    return is_eq


def calc_q_budget(df, header, row_idx):
    temp_q1 = []
    temp_q2 = []
    temp_q3 = []
    temp_q4 = []
    for col_idx in range(3):
        try:
            temp_q1.append(float(df[header[col_idx + 20]][row_idx]))
        except ValueError:
            temp_q1.append(0.0)
    for col_idx in range(3):
        try:
            temp_q2.append(float(df[header[col_idx + 23]][row_idx]))
        except ValueError:
            temp_q2.append(0.0)
    for col_idx in range(3):
        try:
            temp_q3.append(float(df[header[col_idx + 26]][row_idx]))
        except ValueError:
            temp_q3.append(0.0)
    for col_idx in range(3):
        try:
            temp_q4.append(float(df[header[col_idx + 29]][row_idx]))
        except ValueError:
            temp_q4.append(0.0)
    total_q1_budget = sum(list(map(float, temp_q1)))
    total_q2_budget = sum(list(map(float, temp_q2)))
    total_q3_budget = sum(list(map(float, temp_q3)))
    total_q4_budget = sum(list(map(float, temp_q4)))
    return [total_q1_budget, total_q2_budget, total_q3_budget, total_q4_budget]


def validate_team(df, header, row_idx):
    validation = False
    if str(df[header[1]][row_idx]).replace(' ', '').upper() in read_dept(dept_std):
        if str(df[header[2]][row_idx]).replace(' ', '').upper() \
                in read_sub_std(dept_std, df[header[1]][row_idx]):
            validation = True
    return validation


def validate(df, header, row_idx):
    validation_sale = False
    validation_cate = False
    validation_acti = False
    std_working = read_std(working_std)  # ['Awareness', 'Consideration', 'Opinion']
    if str(df[header[6]][row_idx]).replace(' ', '').upper() == 'NONWORKING' \
            and str(df[header[7]][row_idx]).replace(' ', '').upper() not in std_working:
        df[header[6]][row_idx] = 'NonWorking'
        validation_sale = True
        if str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in read_sub_std(nonworking_std, 'Category'):
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in read_sub_std(nonworking_std, df[header[8]][row_idx]):
                validation_acti = True
    elif str(df[header[6]][row_idx]).replace(' ', '').upper() == 'WORKING' \
            and str(df[header[7]][row_idx]).replace(' ', '').upper() in std_working:
        validation_sale = True
        if str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in read_sub_std(working_std, 'Category'):
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in read_sub_std(working_std, df[header[8]][row_idx]):
                validation_acti = True
    return [validation_sale, validation_cate, validation_acti]


def validate_kpi(df, header, row_idx):
    validation = True
    if str(df[header[16]][row_idx]).strip().upper() == 'Y' \
            and str(df[header[11]][row_idx]) in ['0', 'nan', '/', '-'] \
            and str(df[header[12]][row_idx]) in ['0', 'nan', '/', '-'] \
            and str(df[header[13]][row_idx]) in ['0', 'nan', '/', '-'] \
            and str(df[header[14]][row_idx]) in ['0', 'nan', '/', '-']:
        validation = False
    if str(df[header[16]][row_idx]).strip().upper() == 'N' \
            and (str(df[header[11]][row_idx]) not in ['0', 'nan', '/', '-']
                 or str(df[header[12]][row_idx]) not in ['0', 'nan', '/', '-']
                 or str(df[header[13]][row_idx]) not in ['0', 'nan', '/', '-']
                 or str(df[header[14]][row_idx]) not in ['0', 'nan', '/', '-']):
        validation = False
    return validation


def is_duplicate(df, row_idx, issue_name, actual_name, budget_name):
    if str(df['Issue'][row_idx]).replace(' ', '').upper() in issue_name \
            or str(df['Issue'][row_idx]).replace(' ', '').upper() in actual_name\
            or str(df['Issue'][row_idx]).replace(' ', '').upper() in budget_name:
        return True
    else:
        return False


def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
    return False


def is_num_kpi(df, header, row_idx):
    flag_prospects = True
    flag_leads = True
    flag_inquiry = True
    flag_order = True
    if not is_number(str(df[header[11]][row_idx])):
        flag_prospects = False
    if not is_number(str(df[header[12]][row_idx])):
        flag_leads = False
    if not is_number(str(df[header[13]][row_idx])):
        flag_inquiry = False
    if not is_number(str(df[header[14]][row_idx])):
        flag_order = False
    return [flag_prospects, flag_leads, flag_inquiry, flag_order]


def to_float(df, header, col_idx, row_idx):
    if 20 <= col_idx <= 32:
        if not is_number(str(df[header[col_idx]][row_idx])):
            df[header[col_idx]][row_idx] = float(0.0)
        try:
            df[header[col_idx]][row_idx] = float(df[header[col_idx]][row_idx])
        except ValueError:
            df[header[col_idx]][row_idx] = float(0.0)


def to_int(df, header, col_idx, row_idx):
    if 11 <= col_idx <= 14:
        if is_number(str(df[header[col_idx]][row_idx])):
            try:
                df[header[col_idx]][row_idx] = int(df[header[col_idx]][row_idx])
            except ValueError:
                df[header[col_idx]][row_idx] = float(df[header[col_idx]][row_idx])


def to_replace(df, header, col_idx, row_idx):
    new_str = ''
    if 0 <= col_idx <= 10:
        if str(df[header[col_idx]][row_idx]).find('_') != -1:
            df[header[col_idx]][row_idx] = str(df[header[col_idx]][row_idx]).replace('_', ' ')
        if str(df[header[col_idx]][row_idx]).find('-') != -1:
            df[header[col_idx]][row_idx] = str(df[header[col_idx]][row_idx]).replace('-', ' ')
        if col_idx != 5:
            if not str(df[header[col_idx]][row_idx]).isupper():
                if str(df[header[col_idx]][row_idx]).strip().find(' ') == -1:
                    for str_idx in range(len(str(df[header[col_idx]][row_idx]))):
                        if 65 <= ord(str(df[header[col_idx]][row_idx])[str_idx]) <= 90:
                            if str_idx != 0 and 97 <= ord(str(df[header[col_idx]][row_idx])[str_idx - 1]) <= 122:
                                new_str = new_str + ' '
                            new_str = new_str + str(df[header[col_idx]][row_idx])[str_idx]
                        else:
                            new_str = new_str + str(df[header[col_idx]][row_idx])[str_idx]
                    df[header[col_idx]][row_idx] = new_str.strip()


def to_replace_null(df, header, col_idx, row_idx):
    if col_idx in [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15, 16, 18, 19, 33]:
        if str(df[header[col_idx]][row_idx]).replace(' ', '') in ["0", "nan", '-', '/', '']:
            df[header[col_idx]][row_idx] = "N.A."
    if col_idx in [11, 12, 13, 14, 17, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32]:
        if str(df[header[col_idx]][row_idx]).replace(' ', '') in ["0", "nan", '-', '/', '']:
            df[header[col_idx]][row_idx] = int(0)


def is_actual(df, header, row_idx, actual_name):
    flag = False
    if str(df[header[0]][row_idx]).replace(' ', '').upper() in actual_name:
        flag = True
    return flag


def is_forecast(df, header, row_idx, issue_name):
    flag = False
    if str(df[header[0]][row_idx]).replace(' ', '').upper() in issue_name:
        flag = True
    return flag


def is_budget(df, header, row_idx, budget_name):
    flag = False
    if str(df[header[0]][row_idx]).replace(' ', '').upper() in budget_name:
        flag = True
    return flag


def is_actual_only(df, header, row_idx, actual_name):
    flag = False
    if str(df[header[0]][row_idx]).replace(' ', '').upper() in actual_name:
        flag = True
    return flag


def is_release_gap_task(df, header, row_idx):
    flag = False
    if str(df[header[1]][row_idx]).replace(' ', '').upper() in ['RELEASE', 'GAP', 'TASK']:
        flag = True
    return flag


def to_simplify(df, header, row_idx):
    df[header[10]][row_idx] = str(df[header[10]][row_idx]).replace(str(df[header[3]][row_idx]).strip(), '').strip()


def to_upper(df, header, row_idx, col, val):
    if str(df[header[col]][row_idx]).strip().upper() == val.upper():
        df[header[col]][row_idx] = val


def to_format(writer, data_frame, data, err_col, red_col, gray_col):
    # Indicate workbook and worksheet for formatting
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    worksheet.set_row(row=0, height=57)
    # Add a header format.
    header_format = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_color': '#FFFFFF',
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#0070C0',
        'border': 1})
    header_format_err = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_color': '#FFFFFF',
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#FF8000',
        'border': 1})
    header_format_pqt = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_color': '#FFFFFF',
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': 'red',
        'border': 1})
    header_format_rs = workbook.add_format({
        'bold': True,
        'font_name': 'Calibri',
        'font_color': '#FFFFFF',
        'text_wrap': True,
        'align': 'center',
        'valign': 'vcenter',
        'bg_color': '#595959',
        'border': 1})
    budget_format = workbook.add_format({
        'num_format': '#,##0.00'})
    # Write the column headers with the defined format.
    for col_num, value in enumerate(data_frame.columns.values):
        if col_num in err_col:
            worksheet.write(0, col_num, value, header_format_err)
        elif col_num in red_col:
            worksheet.write(0, col_num, value, header_format_pqt)
        elif col_num in gray_col:
            worksheet.write(0, col_num, value, header_format_rs)
        else:
            worksheet.write(0, col_num, value, header_format)
    # Iterate through each column and set the width == the max length in that column.
    # A padding length of 2 is also added.
    for i, col in enumerate(data_frame.columns):
        # find length of column i
        if col in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
                   'Total Budget', 'Q1', 'Q2', 'Q3', 'Q4']:
            column_len = data_frame[col].astype(str).str.len().max()
        else:
            column_len = data_frame[col].astype(str).str.len().mean()
        # Setting the length if the column header is larger than the max column value length
        column_len = max(column_len, len(col)) + 2
        # set the column length
        worksheet.set_column(i, i, column_len)
    try:
        if len(data_frame.columns) == 34:
            for row_idx in range(len(data)):
                for col_idx in range(34):
                    if 20 <= col_idx <= 32:
                        worksheet.write(row_idx + 1, col_idx, data[row_idx][col_idx], budget_format)
        elif len(data_frame.columns) == 37:
            for row_idx in range(len(data)):
                for col_idx in range(37):
                    if 23 <= col_idx <= 35:
                        worksheet.write(row_idx + 1, col_idx, data[row_idx][col_idx], budget_format)
        elif len(data_frame.columns) == 38:
            for row_idx in range(len(data)):
                for col_idx in range(38):
                    if 20 <= col_idx <= 32 or 34 <= col_idx <= 37:
                        worksheet.write(row_idx + 1, col_idx, data[row_idx][col_idx], budget_format)
    except IndexError:
        print("hello")
    writer.save()
    writer.close()


def add_error_log_and_data(f_name, df, header, row_err_data, err_data, exception, row_idx):
    # 添加error log（三列）
    row_err_data.append(f_name)
    row_err_data.append(exception)
    row_err_data.append(row_idx + 2)
    # 添加error data（header列）
    for col_idx in range(len(header)):
        row_err_data.append(df[header[col_idx]][row_idx])
    err_data.append(row_err_data)


def write_to_excel(data, header, writer, err_col, red_col, gray_col):
    frame = pd.DataFrame(data, columns=header)
    frame.to_excel(writer, index=False, header=True)
    to_format(writer, frame, data, err_col, red_col, gray_col)


start = time.time()
# 标准数据excel header
header_std = ['Issue', 'Dept', 'Team', 'Carline', 'Lifecycle', 'Branding/NonBranding',
              'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type',
              'Activity', 'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order',
              'KPI Others', 'SMM Campaign Code (Y/N)', 'SC NO.', ' SC Name',
              'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
              'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
header_std_plus = ['Issue', 'Dept', 'Team', 'Carline', 'Lifecycle', 'Branding/NonBranding',
                   'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type',
                   'Activity', 'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order',
                   'KPI Others', 'SMM Campaign Code (Y/N)', 'SC NO.', ' SC Name',
                   'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
                   'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)', 'Q1', 'Q2', 'Q3', 'Q4']
# 异常数据excel header
header_err = ['File Name', 'Exception Type', 'Index', 'Issue', 'Dept', 'Team', 'Carline', 'Lifecycle',
              'Branding/NonBranding', 'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type', 'Activity',
              'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order', 'KPI Others', 'SMM Campaign Code (Y/N)',
              'SC NO.', ' SC Name', 'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct',
              'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
# 异常header excel header
err_header_err = ['File Name', 'Exception Type']
file_std = "D:/ASP - Erin/Standard.xlsx"
df_folder_path = pd.read_excel(file_std, sheet_name='Folder Path', dtype=str)
dept_std = pd.read_excel(file_std, sheet_name='Department Standard', dtype=str)
carline_std = pd.read_excel(file_std, sheet_name='Carline Standard', dtype=str)
lifecycle_std = pd.read_excel(file_std, sheet_name='Lifecycle Standard', dtype=str)
nonworking_std = pd.read_excel(file_std, sheet_name='NonWorking Standard', dtype=str)
working_std = pd.read_excel(file_std, sheet_name='Working Standard', dtype=str)
df_issue = pd.read_excel(file_std, sheet_name='Issue Standard', dtype=str)
# 读入待处理文件
issue_fcst_list = []
for issue_fcst in df_issue['Issue FCST']:
    if str(issue_fcst) != 'nan':
        issue_fcst_list.append(issue_fcst.replace(' ', '').upper())
issue_actual_list = []
for issue_actual in df_issue['Issue Actual']:
    if str(issue_actual) != 'nan':
        issue_actual_list.append(issue_actual.replace(' ', '').upper())
issue_budget_list = []
for issue_budget in df_issue['Issue Budget']:
    if str(issue_budget) != 'nan':
        issue_budget_list.append(issue_budget.replace(' ', '').upper())

folder_path = df_folder_path['Path'][0]
path = folder_path + "Raw Data/"
file_name_all = os.listdir(path)
file_name = []
for file_idx in range(len(file_name_all)):
    if file_name_all[file_idx].endswith('.xlsx'):
        file_name.append(path + file_name_all[file_idx])
data_new_actual = []  # 存储ACT所有的标准数据
data_new_fcst = []  # 存储FCST所有的标准数据
data_new_bgt = []  # 存储FCST所有的标准数据
data_err_actual = []  # 存储ACT所有的异常数据
data_err_fcst = []  # 存储FCST所有的异常数据
data_err_bgt = []  # 存储FCST所有的异常数据
data_err_header = []  # 存储所有的header异常数据

data_MKT = []  # 存储所有的标准数据
data_MKT_Launch = []
data_APAC_CC = []
data_Customer_Service = []
data_DTS = []
data_MI = []
data_Sales_MKT1 = []
data_Sales_MKT2 = []
data_Sales_MKT3 = []
data_Sales_MKT4 = []
data_Sales_MKT5 = []
# 创建存放excel的文件夹
file_path_a = folder_path + "FCST/Data&Log/"
file_path_b = folder_path + "Header Error/"
file_path_c = folder_path + "Header Error/Files/"
file_path_d = folder_path + "Actual/Data&Log/"
file_path_e = folder_path + "Dept-Team/"
file_path_f = folder_path + "Budget/Data&Log/"
mkdir(file_path_a)
mkdir(file_path_b)
mkdir(file_path_c)
mkdir(file_path_d)
mkdir(file_path_e)
mkdir(file_path_f)
# 几个生成的excel路径及文件名
file_new_actual = folder_path + "Actual/Data&Log/Act_Summary.xlsx"
file_new_fcst = folder_path + "FCST/Data&Log/FCST_Summary.xlsx"
file_new_bgt = folder_path + "Budget/Data&Log/Budget_Summary.xlsx"
file_err_actual = folder_path + "Actual/Data&Log/Act_Error.xlsx"
file_err_fcst = folder_path + "FCST/Data&Log/FCST_Error.xlsx"
file_err_bgt = folder_path + "Budget/Data&Log/Budget_Error.xlsx"
file_header_err = folder_path + "Header Error/Header_Error.xlsx"

file_MKT = folder_path + "Dept-Team/MKT.xlsx"
file_MKT_Launch = folder_path + "Dept-Team/MKT Launch.xlsx"
file_APAC_CC = folder_path + "Dept-Team/APAC CC.xlsx"
file_Customer_Service = folder_path + "Dept-Team/Customer Service.xlsx"
file_DTS = folder_path + "Dept-Team/DTS.xlsx"
file_MI = folder_path + "Dept-Team/MI.xlsx"
file_Sales_MKT1 = folder_path + "Dept-Team/Sales MKT_DMKT.xlsx"
file_Sales_MKT2 = folder_path + "Dept-Team/Sales MKT_HK.xlsx"
file_Sales_MKT3 = folder_path + "Dept-Team/Sales MKT_Internal Fleet Car.xlsx"
file_Sales_MKT4 = folder_path + "Dept-Team/Sales MKT_National Sales.xlsx"
file_Sales_MKT5 = folder_path + "Dept-Team/Sales MKT_NBD Team.xlsx"

# 创建几个excel的Writer引擎
writer_new_actual = pd.ExcelWriter(path=file_new_actual, mode='w', engine='xlsxwriter')
writer_new_fcst = pd.ExcelWriter(path=file_new_fcst, mode='w', engine='xlsxwriter')
writer_new_bgt = pd.ExcelWriter(path=file_new_bgt, mode='w', engine='xlsxwriter')
writer_err_actual = pd.ExcelWriter(path=file_err_actual, mode='w', engine='xlsxwriter')
writer_err_fcst = pd.ExcelWriter(path=file_err_fcst, mode='w', engine='xlsxwriter')
writer_err_bgt = pd.ExcelWriter(path=file_err_bgt, mode='w', engine='xlsxwriter')
writer_err_header = pd.ExcelWriter(path=file_header_err, mode='w', engine='xlsxwriter')
writer_MKT = pd.ExcelWriter(path=file_MKT, mode='w', engine='xlsxwriter')
writer_MKT_Launch = pd.ExcelWriter(path=file_MKT_Launch, mode='w', engine='xlsxwriter')
writer_APAC_CC = pd.ExcelWriter(path=file_APAC_CC, mode='w', engine='xlsxwriter')
writer_Customer_Service = pd.ExcelWriter(path=file_Customer_Service, mode='w', engine='xlsxwriter')
writer_DTS = pd.ExcelWriter(path=file_DTS, mode='w', engine='xlsxwriter')
writer_MI = pd.ExcelWriter(path=file_MI, mode='w', engine='xlsxwriter')
writer_Sales_MKT1 = pd.ExcelWriter(path=file_Sales_MKT1, mode='w', engine='xlsxwriter')
writer_Sales_MKT2 = pd.ExcelWriter(path=file_Sales_MKT2, mode='w', engine='xlsxwriter')
writer_Sales_MKT3 = pd.ExcelWriter(path=file_Sales_MKT3, mode='w', engine='xlsxwriter')
writer_Sales_MKT4 = pd.ExcelWriter(path=file_Sales_MKT4, mode='w', engine='xlsxwriter')
writer_Sales_MKT5 = pd.ExcelWriter(path=file_Sales_MKT5, mode='w', engine='xlsxwriter')
# 从文件列表遍历读取文件
for file_idx in range(len(file_name)):
    df_asp = pd.read_excel(file_name[file_idx], dtype=str)
    # 判断header是否符合标准
    header_result = str(list(df_asp.columns[0:34])).strip() == str(header_std)
    if not header_result:
        row_data_err = [file_name[file_idx], 'Header Exception']
        data_err_header.append(row_data_err)
        try:
            shutil.copy(file_name[file_idx], file_path_c)
        except shutil.SameFileError:
            pass
    else:
        for idx_row in range(len(df_asp[header_std[0]])):
            row_data_MKT = []
            row_data_MKT_Launch = []
            row_data_APAC_CC = []
            row_data_Customer_Service = []
            row_data_DTS = []
            row_data_MI = []
            row_data_Sales_MKT1 = []
            row_data_Sales_MKT2 = []
            row_data_Sales_MKT3 = []
            row_data_Sales_MKT4 = []
            row_data_Sales_MKT5 = []
            # data clean
            for idx_col in range(len(header_std)):
                to_replace_null(df_asp, header_std, idx_col, idx_row)
                to_replace(df_asp, header_std, idx_col, idx_row)
                to_simplify(df_asp, header_std, idx_row)
                to_upper(df_asp, header_std, idx_row, 5, 'Branding')
                to_upper(df_asp, header_std, idx_row, 5, 'NonBranding')
                to_upper(df_asp, header_std, idx_row, 6, 'Working')
                to_upper(df_asp, header_std, idx_row, 6, 'NonWorking')
                to_float(df_asp, header_std, idx_col, idx_row)
                to_int(df_asp, header_std, idx_col, idx_row)

            if str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKT':
                for idx_col in range(len(header_std)):
                    row_data_MKT.append(df_asp[header_std[idx_col]][idx_row])
                data_MKT.append(row_data_MKT)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKTLAUNCH':
                for idx_col in range(len(header_std)):
                    row_data_MKT_Launch.append(df_asp[header_std[idx_col]][idx_row])
                data_MKT_Launch.append(row_data_MKT_Launch)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'APACCC':
                for idx_col in range(len(header_std)):
                    row_data_APAC_CC.append(df_asp[header_std[idx_col]][idx_row])
                data_APAC_CC.append(row_data_APAC_CC)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'CUSTOMERSERVICE':
                for idx_col in range(len(header_std)):
                    row_data_Customer_Service.append(df_asp[header_std[idx_col]][idx_row])
                data_Customer_Service.append(row_data_Customer_Service)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'DTS':
                for idx_col in range(len(header_std)):
                    to_float(df_asp, header_std, idx_col, idx_row)
                    row_data_DTS.append(df_asp[header_std[idx_col]][idx_row])
                data_DTS.append(row_data_DTS)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MI':
                for idx_col in range(len(header_std)):
                    row_data_MI.append(df_asp[header_std[idx_col]][idx_row])
                data_MI.append(row_data_MI)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] \
                    and str(df_asp[header_std[2]][idx_row]).replace(' ', '').upper() \
                    in ['DMKTCENTRALTEAM', 'EXHIBITION', 'EASTREGIONTEAM', 'NORTHREGIONTEAM', 'SOUTHREGIONTEAM',
                        'WESTREGIONTEAM', 'ZHEJIANGREGIONTEAM']:
                for idx_col in range(len(header_std)):
                    row_data_Sales_MKT1.append(df_asp[header_std[idx_col]][idx_row])
                data_Sales_MKT1.append(row_data_Sales_MKT1)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                    df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['HK']:
                for idx_col in range(len(header_std)):
                    row_data_Sales_MKT2.append(df_asp[header_std[idx_col]][idx_row])
                data_Sales_MKT2.append(row_data_Sales_MKT2)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                    df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['INTERNALFLEETCAR']:
                for idx_col in range(len(header_std)):
                    row_data_Sales_MKT3.append(df_asp[header_std[idx_col]][idx_row])
                data_Sales_MKT3.append(row_data_Sales_MKT3)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                    df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['NATIONALSALES']:
                for idx_col in range(len(header_std)):
                    row_data_Sales_MKT4.append(df_asp[header_std[idx_col]][idx_row])
                data_Sales_MKT4.append(row_data_Sales_MKT4)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] \
                    and str(df_asp[header_std[2]][idx_row]).replace(' ', '').upper() \
                    in ['FLEETTEAM', 'USEDCARTEAM', 'FINANCIALSERVICETEAM']:
                for idx_col in range(len(header_std)):
                    row_data_Sales_MKT5.append(df_asp[header_std[idx_col]][idx_row])
                data_Sales_MKT5.append(row_data_Sales_MKT5)
            # 符合header标准，遍历文件的所有行，验证行数据 data checking
            if is_actual(df_asp, header_std, idx_row, issue_actual_list):
                row_data = []  # 临时存放符合标准的idx_row行的数据
                row_data_err_null = []  # 临时存放错误的idx_row行的数据
                row_data_err_calc = []  # 临时存放错误的idx_row行的数据
                row_data_err_dupl = []  # 临时存放错误的idx_row行的数据
                row_data_err_vali = []  # 临时存放错误的idx_row行的数据
                row_data_err_ssmc = []  # 临时存放错误的idx_row行的数据
                row_data_err_kpi = []  # 临时存放错误的idx_row行的数据
                row_data_err_brand = []  # 临时存放错误的idx_row行的数据
                row_data_err_team = []  # 临时存放错误的idx_row行的数据
                row_data_err_carl = []  # 临时存放错误的idx_row行的数据
                row_data_err_life = []  # 临时存放错误的idx_row行的数据
                row_data_err_kpin = []  # 临时存放错误的idx_row行的数据
                if is_release_gap_task(df_asp, header_std, idx_row):
                    for idx_col in range(len(header_std)):
                        row_data.append(df_asp[header_std[idx_col]][idx_row])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[0])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[1])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[2])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[3])
                    # row_data.append(calc_ytd_budget(df_asp, header_std, idx_row, cur_month))
                    data_new_actual.append(row_data)
                else:
                    # 判断idx_row的前两列是否为空
                    if not is_null(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_null, data_err_actual,
                                               'Issue/Dept Null Exception', idx_row)
                    if not is_carline(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_carl, data_err_actual,
                                               'Carline Exception', idx_row)
                    if not is_life_cycle(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_life, data_err_actual,
                                               'Lifecycle Exception', idx_row)
                    # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                    if not calc_total_budget(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_calc, data_err_actual,
                                               'Budget Exception', idx_row)
                    if not is_branding(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_brand, data_err_actual,
                                               'Branding/NonBranding Exception', idx_row)
                    if not is_duplicate(df_asp, idx_row, issue_fcst_list, issue_actual_list, issue_budget_list):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_dupl, data_err_actual,
                                               'Issue Exception', idx_row)
                    # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
                    if not is_num_kpi(df_asp, header_std, idx_row)[0]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin, data_err_actual,
                                               'KPI Prospects Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[1]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin, data_err_actual,
                                               'KPI Leads Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[2]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin, data_err_actual,
                                               'KPI Inquiry Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[3]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin, data_err_actual,
                                               'KPI Order Exception', idx_row)
                    if not validate_team(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_team, data_err_actual,
                                               'Team Exception', idx_row)
                    if not validate(df_asp, header_std, idx_row)[0]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_vali, data_err_actual,
                                               'Sale Funnel Exception', idx_row)
                    elif not validate(df_asp, header_std, idx_row)[1]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_vali, data_err_actual,
                                               'Category Exception', idx_row)
                    elif not validate(df_asp, header_std, idx_row)[2]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_vali, data_err_actual,
                                               'Activity Type Exception', idx_row)
                    if not is_ssm(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_ssmc, data_err_actual,
                                               'SMM Exception', idx_row)
                    if not validate_kpi(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpi, data_err_actual,
                                               'KPI Exception', idx_row)
                    # 符合所有标准的数据格式化后写入Summary文件
                    if is_null(df_asp, header_std, idx_row) and calc_total_budget(df_asp, header_std, idx_row) \
                            and is_branding(df_asp, header_std, idx_row) and validate(df_asp, header_std, idx_row)[0] \
                            and validate(df_asp, header_std, idx_row)[1] and validate(df_asp, header_std, idx_row)[2] \
                            and is_ssm(df_asp, header_std, idx_row) and validate_kpi(df_asp, header_std, idx_row) \
                            and is_carline(df_asp, header_std, idx_row) and is_life_cycle(df_asp, header_std, idx_row) \
                            and validate_team(df_asp, header_std, idx_row) and is_num_kpi(df_asp, header_std, idx_row)[0] \
                            and is_num_kpi(df_asp, header_std, idx_row)[1] and is_num_kpi(df_asp, header_std, idx_row)[2] \
                            and is_num_kpi(df_asp, header_std, idx_row)[3] \
                            and is_duplicate(df_asp, idx_row, issue_fcst_list, issue_actual_list, issue_budget_list):
                        for idx_col in range(len(header_std)):
                            row_data.append(df_asp[header_std[idx_col]][idx_row])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[0])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[1])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[2])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[3])
                        # row_data.append(calc_ytd_budget(df_asp, header_std, idx_row, cur_month))
                        data_new_actual.append(row_data)
            if is_forecast(df_asp, header_std, idx_row, issue_fcst_list):
                row_data = []  # 临时存放符合标准的idx_row行的数据
                row_data_err_null_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_calc_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_dupl_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_vali_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_ssmc_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_kpi_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_brand_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_team_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_carl_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_life_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_kpin_fcst = []  # 临时存放错误的idx_row行的数据
                if is_release_gap_task(df_asp, header_std, idx_row):
                    for idx_col in range(len(header_std)):
                        row_data.append(df_asp[header_std[idx_col]][idx_row])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[0])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[1])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[2])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[3])
                    # row_data.append(calc_ytd_budget(df_asp, header_std, idx_row, cur_month))
                    data_new_fcst.append(row_data)
                else:
                    # 判断idx_row的前两列是否为空
                    if not is_null(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_null_fcst,
                                               data_err_fcst,
                                               'Issue/Dept Null Exception', idx_row)
                    if not is_carline(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_carl_fcst,
                                               data_err_fcst,
                                               'Carline Exception', idx_row)
                    if not is_life_cycle(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_life_fcst,
                                               data_err_fcst,
                                               'Lifecycle Exception', idx_row)
                    # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                    if not calc_total_budget(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_calc_fcst,
                                               data_err_fcst,
                                               'TotalBudget Exception', idx_row)
                    if not is_branding(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_brand_fcst,
                                               data_err_fcst,
                                               'Branding/NonBranding Exception', idx_row)
                    if not is_num_kpi(df_asp, header_std, idx_row)[0]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin_fcst,
                                               data_err_fcst,
                                               'KPI Prospects Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[1]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin_fcst,
                                               data_err_fcst,
                                               'KPI Leads Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[2]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin_fcst,
                                               data_err_fcst,
                                               'KPI Inquiry Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[3]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin_fcst,
                                               data_err_fcst,
                                               'KPI Order Exception', idx_row)
                    if not is_duplicate(df_asp, idx_row, issue_fcst_list, issue_actual_list, issue_budget_list):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_dupl_fcst,
                                               data_err_fcst,
                                               'Issue Exception', idx_row)
                    # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
                    if not validate_team(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_team_fcst,
                                               data_err_fcst,
                                               'Team Exception', idx_row)
                    if not validate(df_asp, header_std, idx_row)[0]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_vali_fcst,
                                               data_err_fcst,
                                               'Sale Funnel Exception', idx_row)
                    elif not validate(df_asp, header_std, idx_row)[1]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_vali_fcst,
                                               data_err_fcst,
                                               'Category Exception', idx_row)
                    elif not validate(df_asp, header_std, idx_row)[2]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_vali_fcst,
                                               data_err_fcst,
                                               'Activity Type Exception', idx_row)
                    if not is_ssm(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_ssmc_fcst,
                                               data_err_fcst,
                                               'SMM Exception', idx_row)
                    if not validate_kpi(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpi_fcst,
                                               data_err_fcst,
                                               'KPI Exception', idx_row)
                    # 符合所有标准的数据格式化后写入Summary文件
                    if is_null(df_asp, header_std, idx_row) and calc_total_budget(df_asp, header_std, idx_row) \
                            and is_branding(df_asp, header_std, idx_row) and validate(df_asp, header_std, idx_row)[0] \
                            and validate(df_asp, header_std, idx_row)[1] and validate(df_asp, header_std, idx_row)[2] \
                            and is_ssm(df_asp, header_std, idx_row) and validate_kpi(df_asp, header_std, idx_row) \
                            and is_duplicate(df_asp, idx_row, issue_fcst_list, issue_actual_list, issue_budget_list) \
                            and is_carline(df_asp, header_std, idx_row) and is_life_cycle(df_asp, header_std, idx_row) \
                            and validate_team(df_asp, header_std, idx_row) and is_num_kpi(df_asp, header_std, idx_row)[0] \
                            and is_num_kpi(df_asp, header_std, idx_row)[1] and is_num_kpi(df_asp, header_std, idx_row)[2] \
                            and is_num_kpi(df_asp, header_std, idx_row)[3]:
                        for idx_col in range(len(header_std)):
                            row_data.append(df_asp[header_std[idx_col]][idx_row])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[0])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[1])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[2])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[3])
                        # row_data.append(calc_ytd_budget(df_asp, header_std, idx_row, cur_month))
                        data_new_fcst.append(row_data)
            if is_budget(df_asp, header_std, idx_row, issue_budget_list):
                row_data = []  # 临时存放符合标准的idx_row行的数据
                row_data_err_null_bgt = []  # 临时存放错误的idx_row行的数据
                row_data_err_calc_bgt = []  # 临时存放错误的idx_row行的数据
                row_data_err_dupl_bgt = []  # 临时存放错误的idx_row行的数据
                row_data_err_vali_bgt = []  # 临时存放错误的idx_row行的数据
                row_data_err_ssmc_bgt = []  # 临时存放错误的idx_row行的数据
                row_data_err_kpi_bgt = []  # 临时存放错误的idx_row行的数据
                row_data_err_brand_bgt = []  # 临时存放错误的idx_row行的数据
                row_data_err_team_bgt = []  # 临时存放错误的idx_row行的数据
                row_data_err_carl_bgt = []  # 临时存放错误的idx_row行的数据
                row_data_err_life_bgt = []  # 临时存放错误的idx_row行的数据
                row_data_err_kpin_bgt = []  # 临时存放错误的idx_row行的数据
                if is_release_gap_task(df_asp, header_std, idx_row):
                    for idx_col in range(len(header_std)):
                        row_data.append(df_asp[header_std[idx_col]][idx_row])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[0])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[1])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[2])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[3])
                    # row_data.append(calc_ytd_budget(df_asp, header_std, idx_row, cur_month))
                    data_new_bgt.append(row_data)
                else:
                    # 判断idx_row的前两列是否为空
                    if not is_null(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_null_bgt,
                                               data_err_bgt,
                                               'Issue/Dept Null Exception', idx_row)
                    if not is_carline(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_carl_bgt,
                                               data_err_bgt,
                                               'Carline Exception', idx_row)
                    if not is_life_cycle(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_life_bgt,
                                               data_err_bgt,
                                               'Lifecycle Exception', idx_row)
                    # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                    if not calc_total_budget(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_calc_bgt,
                                               data_err_bgt,
                                               'TotalBudget Exception', idx_row)
                    if not is_branding(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_brand_bgt,
                                               data_err_bgt,
                                               'Branding/NonBranding Exception', idx_row)
                    if not is_num_kpi(df_asp, header_std, idx_row)[0]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin_bgt,
                                               data_err_bgt,
                                               'KPI Prospects Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[1]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin_bgt,
                                               data_err_bgt,
                                               'KPI Leads Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[2]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin_bgt,
                                               data_err_bgt,
                                               'KPI Inquiry Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[3]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpin_bgt,
                                               data_err_bgt,
                                               'KPI Order Exception', idx_row)
                    if not is_duplicate(df_asp, idx_row, issue_fcst_list, issue_actual_list, issue_budget_list):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_dupl_bgt,
                                               data_err_bgt,
                                               'Issue Exception', idx_row)
                    # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
                    if not validate_team(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_team_bgt,
                                               data_err_bgt,
                                               'Team Exception', idx_row)
                    if not validate(df_asp, header_std, idx_row)[0]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_vali_bgt,
                                               data_err_bgt,
                                               'Sale Funnel Exception', idx_row)
                    elif not validate(df_asp, header_std, idx_row)[1]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_vali_bgt,
                                               data_err_bgt,
                                               'Category Exception', idx_row)
                    elif not validate(df_asp, header_std, idx_row)[2]:
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_vali_bgt,
                                               data_err_bgt,
                                               'Activity Type Exception', idx_row)
                    if not is_ssm(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_ssmc_bgt,
                                               data_err_bgt,
                                               'SMM Exception', idx_row)
                    if not validate_kpi(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err_kpi_bgt,
                                               data_err_bgt,
                                               'KPI Exception', idx_row)
                    # 符合所有标准的数据格式化后写入Summary文件
                    if is_null(df_asp, header_std, idx_row) and calc_total_budget(df_asp, header_std, idx_row) \
                            and is_branding(df_asp, header_std, idx_row) and validate(df_asp, header_std, idx_row)[0] \
                            and validate(df_asp, header_std, idx_row)[1] and validate(df_asp, header_std, idx_row)[2] \
                            and is_ssm(df_asp, header_std, idx_row) and validate_kpi(df_asp, header_std, idx_row) \
                            and is_duplicate(df_asp, idx_row, issue_fcst_list, issue_actual_list, issue_budget_list) \
                            and is_carline(df_asp, header_std, idx_row) and is_life_cycle(df_asp, header_std, idx_row) \
                            and validate_team(df_asp, header_std, idx_row) and is_num_kpi(df_asp, header_std, idx_row)[0] \
                            and is_num_kpi(df_asp, header_std, idx_row)[1] and is_num_kpi(df_asp, header_std, idx_row)[2] \
                            and is_num_kpi(df_asp, header_std, idx_row)[3]:
                        for idx_col in range(len(header_std)):
                            row_data.append(df_asp[header_std[idx_col]][idx_row])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[0])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[1])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[2])
                        row_data.append(calc_q_budget(df_asp, header_std, idx_row)[3])
                        # row_data.append(calc_ytd_budget(df_asp, header_std, idx_row, cur_month))
                        data_new_bgt.append(row_data)

# 将header异常数据写入Error excel
write_to_excel(data_err_header, err_header_err, writer_err_header, [0, 1], [], [])
# 将Actual异常数据写入Error excel
write_to_excel(data_err_actual, header_err, writer_err_actual, [0, 1, 2], [18, 19, 22, 35], [20, 21])
# 将FCST异常数据写入excel
write_to_excel(data_err_fcst, header_err, writer_err_fcst, [0, 1, 2], [18, 19, 22, 35], [20, 21])
# 将Budget异常数据写入excel
write_to_excel(data_err_bgt, header_err, writer_err_bgt, [0, 1, 2], [18, 19, 22, 35], [20, 21])
# 将Actual标准数据写入excel
write_to_excel(data_new_actual, header_std_plus, writer_new_actual, [], [15, 16, 19, 32], [17, 18])
# 将FCST标准数据写入Error excel
write_to_excel(data_new_fcst, header_std_plus, writer_new_fcst, [], [15, 16, 19, 32], [17, 18])
# 将FCST标准数据写入Error excel
write_to_excel(data_new_bgt, header_std_plus, writer_new_bgt, [], [15, 16, 19, 32], [17, 18])

write_to_excel(data_MKT, header_std, writer_MKT, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_MKT_Launch, header_std, writer_MKT_Launch, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_APAC_CC, header_std, writer_APAC_CC, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Customer_Service, header_std, writer_Customer_Service, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_DTS, header_std, writer_DTS, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_MI, header_std, writer_MI, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT1, header_std, writer_Sales_MKT1, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT2, header_std, writer_Sales_MKT2, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT3, header_std, writer_Sales_MKT3, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT4, header_std, writer_Sales_MKT4, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT5, header_std, writer_Sales_MKT5, [], [15, 16, 19, 32], [17, 18])
os.chmod(file_new_actual, stat.S_IREAD)
os.chmod(file_new_fcst, stat.S_IREAD)
os.chmod(file_new_bgt, stat.S_IREAD)
cost = time.time() - start
print(str(int(cost)) + "s")
# os.system("pause")