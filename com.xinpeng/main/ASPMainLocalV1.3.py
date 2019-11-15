import datetime
import time
import pandas as pd
import os
import shutil
import stat
import sys


def mkdir(file_path):
    if not os.path.exists(file_path):
        os.makedirs(file_path)


def change_dir_before_excute(src_path):
    path_list = ['Result/Checking/Actual/Data&Log/', 'Result/Checking/FCST/Data&Log/', 'Result/Checking/Budget/Data&Log/']
    for path_item in path_list:
        temp_path = src_path + path_item
        try:
            os.listdir(temp_path)
        except FileNotFoundError:
            temp_path = src_path + '/' + path_item
        if os.listdir(temp_path):
            for item in os.listdir(temp_path):
                os.chmod(temp_path + item, stat.S_IWRITE)


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


def show_time(start_time):
    print("The program is running and it has been running for %.2fseconds" % (time.time() - start_time))
    # os.system('cls')


def mk_time_dir(time_file_path):
    cur_year = datetime.datetime.now().year
    cur_month = datetime.datetime.now().month
    month_list = []
    for idx_mon in range(7):
        month = cur_month + idx_mon
        month_temp = ''
        if month <= 9:
            month_temp = str(cur_year) + '0' + str(month)
        elif 10 <= month <= 12:
            month_temp = str(cur_year) + str(month)
        elif month > 12:
            month_temp = str(cur_year + 1) + '0' + str(abs(month - 12))
        month_list.append(month_temp)
    for idx_month in month_list:
        time_file_path_temp = time_file_path + str(idx_month)
        if not os.path.exists(time_file_path_temp):
            os.makedirs(time_file_path_temp)
            path_data_check = time_file_path_temp + "/Raw Data/Checking/"
            path_data_parse = time_file_path_temp + "/Raw Data/Parsing/"
            path_result_check = time_file_path_temp + "/Result/Checking/"
            path_result_parse = time_file_path_temp + "/Result/Parsing/"
            mkdir(path_data_check)
            mkdir(path_data_parse)
            mkdir(path_result_check)
            mkdir(path_result_parse)


print("The program is starting ... ...")
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
file_std = "W:/Finance Volvo/BIDataSource/Finance/ASP Database/Standard.xlsx"
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
show_time(start)
folder_path = df_folder_path['Path'][0]

time_folder_path = 'W:/Finance Volvo/BIDataSource/Finance/ASP Database/'
mk_time_dir(time_folder_path)
path = folder_path + "Raw Data/Checking/"
try:
    os.listdir(path)
except FileNotFoundError:
    folder_path = folder_path + '/'
    path = folder_path + "Raw Data/Checking/"
file_name_all = os.listdir(path)
file_name = []
for file_idx in range(len(file_name_all)):
    if file_name_all[file_idx].endswith('.xlsx'):
        file_name.append(path + file_name_all[file_idx])
show_time(start)
data_new_actual = []  # 存储ACT所有的标准数据
data_new_fcst = []  # 存储FCST所有的标准数据
data_new_bgt = []  # 存储FCST所有的标准数据
data_err_actual = []  # 存储ACT所有的异常数据
data_err_fcst = []  # 存储FCST所有的异常数据
data_err_bgt = []  # 存储FCST所有的异常数据
data_err_header = []  # 存储所有的header异常数据
path_result_check = folder_path + "Result/Checking/"
path_result_parse = folder_path + "Result/Parsing/"
mkdir(path_result_check)
mkdir(path_result_parse)
# 创建存放excel的文件夹
file_path_a = folder_path + "Result/Checking/FCST/Data&Log/"
file_path_b = folder_path + "Result/Checking/Header Error/"
file_path_c = folder_path + "Result/Checking/Header Error/Files/"
file_path_d = folder_path + "Result/Checking/Actual/Data&Log/"
file_path_f = folder_path + "Result/Checking/Budget/Data&Log/"
mkdir(file_path_a)
mkdir(file_path_b)
mkdir(file_path_c)
mkdir(file_path_d)
mkdir(file_path_f)
show_time(start)

change_dir_before_excute(folder_path)
# 几个生成的excel路径及文件名
file_new_actual = folder_path + "Result/Checking/Actual/Data&Log/Act_Summary.xlsx"
file_new_fcst = folder_path + "Result/Checking/FCST/Data&Log/FCST_Summary.xlsx"
file_new_bgt = folder_path + "Result/Checking/Budget/Data&Log/Budget_Summary.xlsx"
file_err_actual = folder_path + "Result/Checking/Actual/Data&Log/Act_Error.xlsx"
file_err_fcst = folder_path + "Result/Checking/FCST/Data&Log/FCST_Error.xlsx"
file_err_bgt = folder_path + "Result/Checking/Budget/Data&Log/Budget_Error.xlsx"
file_header_err = folder_path + "Result/Checking/Header Error/Header_Error.xlsx"
# 创建几个excel的Writer引擎
writer_new_actual = pd.ExcelWriter(path=file_new_actual, mode='w', engine='xlsxwriter')
writer_new_fcst = pd.ExcelWriter(path=file_new_fcst, mode='w', engine='xlsxwriter')
writer_new_bgt = pd.ExcelWriter(path=file_new_bgt, mode='w', engine='xlsxwriter')
writer_err_actual = pd.ExcelWriter(path=file_err_actual, mode='w', engine='xlsxwriter')
writer_err_fcst = pd.ExcelWriter(path=file_err_fcst, mode='w', engine='xlsxwriter')
writer_err_bgt = pd.ExcelWriter(path=file_err_bgt, mode='w', engine='xlsxwriter')
writer_err_header = pd.ExcelWriter(path=file_header_err, mode='w', engine='xlsxwriter')
# show_time(start)
# 从文件列表遍历读取文件
for file_idx in range(len(file_name)):
    df_asp = pd.read_excel(file_name[file_idx], dtype=str)
    show_time(start)
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
            show_time(start)
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
            if is_actual(df_asp, header_std, idx_row, issue_actual_list):
                row_data = []  # 临时存放符合标准的idx_row行的数据
                row_data_err = []  # 临时存放错误的idx_row行的数据
                row_data_exceptions = ""
                if is_release_gap_task(df_asp, header_std, idx_row):
                    for idx_col in range(len(header_std)):
                        row_data.append(df_asp[header_std[idx_col]][idx_row])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[0])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[1])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[2])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[3])
                    data_new_actual.append(row_data)
                else:
                    # 判断idx_row的前两列是否为空
                    if not is_null(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Issue/Dept Null Exception & '
                    if not is_carline(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Carline Exception & '
                    if not is_life_cycle(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Lifecycle Exception & '
                    # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                    if not calc_total_budget(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Budget Exception & '
                    if not is_branding(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Branding/NonBranding Exception & '
                    if not is_duplicate(df_asp, idx_row, issue_fcst_list, issue_actual_list, issue_budget_list):
                        row_data_exceptions = row_data_exceptions + 'Issue Exception & '
                    # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
                    if not is_num_kpi(df_asp, header_std, idx_row)[0]:
                        row_data_exceptions = row_data_exceptions + 'KPI Prospects Exception & '
                    elif not is_num_kpi(df_asp, header_std, idx_row)[1]:
                        row_data_exceptions = row_data_exceptions + 'KPI Leads Exception & '
                    elif not is_num_kpi(df_asp, header_std, idx_row)[2]:
                        row_data_exceptions = row_data_exceptions + 'KPI Inquiry Exception & '
                    elif not is_num_kpi(df_asp, header_std, idx_row)[3]:
                        row_data_exceptions = row_data_exceptions + 'KPI Order Exception & '
                    if not validate_team(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Team Exception & '
                    if not validate(df_asp, header_std, idx_row)[0]:
                        row_data_exceptions = row_data_exceptions + 'Sale Funnel Exception & '
                    elif not validate(df_asp, header_std, idx_row)[1]:
                        row_data_exceptions = row_data_exceptions + 'Category Exception & '
                    elif not validate(df_asp, header_std, idx_row)[2]:
                        row_data_exceptions = row_data_exceptions + 'Activity Type Exception & '
                    if not is_ssm(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'SMM Exception & '
                    if not validate_kpi(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'KPI Exception & '
                    if row_data_exceptions != "":
                        row_data_exceptions = row_data_exceptions[:-3]
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err, data_err_actual,
                                               row_data_exceptions, idx_row)
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
                row_data_err = []  # 临时存放错误的idx_row行的数据
                row_data_exceptions = ""
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
                        row_data_exceptions = row_data_exceptions + 'Issue/Dept Null Exception & '
                    if not is_carline(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Carline Exception & '
                    if not is_life_cycle(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Lifecycle Exception & '
                    # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                    if not calc_total_budget(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Budget Exception & '
                    if not is_branding(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Branding/NonBranding Exception & '
                    if not is_duplicate(df_asp, idx_row, issue_fcst_list, issue_actual_list, issue_budget_list):
                        row_data_exceptions = row_data_exceptions + 'Issue Exception & '
                    # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
                    if not is_num_kpi(df_asp, header_std, idx_row)[0]:
                        row_data_exceptions = row_data_exceptions + 'KPI Prospects Exception & '
                    elif not is_num_kpi(df_asp, header_std, idx_row)[1]:
                        row_data_exceptions = row_data_exceptions + 'KPI Leads Exception & '
                    elif not is_num_kpi(df_asp, header_std, idx_row)[2]:
                        row_data_exceptions = row_data_exceptions + 'KPI Inquiry Exception & '
                    elif not is_num_kpi(df_asp, header_std, idx_row)[3]:
                        row_data_exceptions = row_data_exceptions + 'KPI Order Exception & '
                    if not validate_team(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Team Exception & '
                    if not validate(df_asp, header_std, idx_row)[0]:
                        row_data_exceptions = row_data_exceptions + 'Sale Funnel Exception & '
                    elif not validate(df_asp, header_std, idx_row)[1]:
                        row_data_exceptions = row_data_exceptions + 'Category Exception & '
                    elif not validate(df_asp, header_std, idx_row)[2]:
                        row_data_exceptions = row_data_exceptions + 'Activity Type Exception & '
                    if not is_ssm(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'SMM Exception & '
                    if not validate_kpi(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'KPI Exception & '
                    if row_data_exceptions != "":
                        row_data_exceptions = row_data_exceptions[:-3]
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err, data_err_fcst,
                                               row_data_exceptions, idx_row)
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
                row_data_err = []  # 临时存放错误的idx_row行的数据
                row_data_exceptions = ""
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
                        row_data_exceptions = row_data_exceptions + 'Issue/Dept Null Exception & '
                    if not is_carline(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Carline Exception & '
                    if not is_life_cycle(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Lifecycle Exception & '
                    # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                    if not calc_total_budget(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Budget Exception & '
                    if not is_branding(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Branding/NonBranding Exception & '
                    if not is_duplicate(df_asp, idx_row, issue_fcst_list, issue_actual_list, issue_budget_list):
                        row_data_exceptions = row_data_exceptions + 'Issue Exception & '
                    # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
                    if not is_num_kpi(df_asp, header_std, idx_row)[0]:
                        row_data_exceptions = row_data_exceptions + 'KPI Prospects Exception & '
                    elif not is_num_kpi(df_asp, header_std, idx_row)[1]:
                        row_data_exceptions = row_data_exceptions + 'KPI Leads Exception & '
                    elif not is_num_kpi(df_asp, header_std, idx_row)[2]:
                        row_data_exceptions = row_data_exceptions + 'KPI Inquiry Exception & '
                    elif not is_num_kpi(df_asp, header_std, idx_row)[3]:
                        row_data_exceptions = row_data_exceptions + 'KPI Order Exception & '
                    if not validate_team(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'Team Exception & '
                    if not validate(df_asp, header_std, idx_row)[0]:
                        row_data_exceptions = row_data_exceptions + 'Sale Funnel Exception & '
                    elif not validate(df_asp, header_std, idx_row)[1]:
                        row_data_exceptions = row_data_exceptions + 'Category Exception & '
                    elif not validate(df_asp, header_std, idx_row)[2]:
                        row_data_exceptions = row_data_exceptions + 'Activity Type Exception & '
                    if not is_ssm(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'SMM Exception & '
                    if not validate_kpi(df_asp, header_std, idx_row):
                        row_data_exceptions = row_data_exceptions + 'KPI Exception & '
                    if row_data_exceptions != "":
                        row_data_exceptions = row_data_exceptions[:-3]
                        add_error_log_and_data(file_name[file_idx], df_asp, header_std, row_data_err, data_err_bgt,
                                               row_data_exceptions, idx_row)
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
show_time(start)
os.chmod(file_new_actual, stat.S_IREAD)
os.chmod(file_new_fcst, stat.S_IREAD)
os.chmod(file_new_bgt, stat.S_IREAD)
os.system('cls')
print("The program has ended. The program runs for a total of %.2fseconds" % (time.time() - start))
os.system("pause")