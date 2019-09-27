import time
import pandas as pd
import os
import shutil
import datetime


def mkdir(file_path):
    if not os.path.exists(file_path):
        os.makedirs(file_path)


def is_null(row_idx):
    is_na = True
    if str(df_fcst[header_std[0]][row_idx]) == "nan" or str(df_fcst[header_std[1]][row_idx]) == "nan":
        is_na = False
    return is_na


def is_ssm(row_idx):
    is_ssmc = True
    if str(df_fcst[header_std[16]][row_idx]) not in ['Y', 'y', 'N', 'n']:
        is_ssmc = False
    return is_ssmc


def calc_total_budget(row_idx):
    temp = []
    for col_idx in range(12):
        if str(df_fcst[header_std[col_idx + 20]][row_idx]) == '':
            df_fcst[header_std[col_idx + 20]][row_idx] = 0.0
            temp.append(0.0)
        elif str(df_fcst[header_std[col_idx + 20]][row_idx]).replace(' ', '') == '':
            df_fcst[header_std[col_idx + 20]][row_idx] = 0.0
            temp.append(0.0)
        elif str(df_fcst[header_std[col_idx + 20]][row_idx]).replace('-', '').replace(' ', '') == '':
            df_fcst[header_std[col_idx + 20]][row_idx] = 0.0
            temp.append(0.0)
        elif str(df_fcst[header_std[col_idx + 20]][row_idx]) == 'nan':
            df_fcst[header_std[col_idx + 20]][row_idx] = 0.0
            temp.append(0.0)
        else:
            temp.append(float(df_fcst[header_std[col_idx + 20]][row_idx]))
    total_budget = sum(list(map(float, temp)))
    is_eq = total_budget == float(df_fcst[header_std[32]][row_idx]) or abs(
        total_budget - float(df_fcst[header_std[32]][row_idx])) <= 1
    return is_eq


def validate(row_idx):
    validation = True
    if str(df_fcst[header_std[6]][row_idx]).upper() not in ['NONWORKING', 'WORKING']:
        validation = False
    elif str(df_fcst[header_std[6]][row_idx]).upper() == 'NONWORKING' and str(df_fcst[header_std[7]][row_idx]) not in \
            ["0", "nan", "N.A."]:
        if df_fcst[header_std[8]][row_idx] not in ['Agency Fee', 'Market Intelligence', 'Marketing Audits',
                                                   'Production', 'Infrastructure Development', 'Meetings',
                                                   'Operational Cost', 'Celebrity']:
            if df_fcst[header_std[9]][row_idx] not in ['Media Agency', 'Event Agency', 'Creative production Agency',
                                                       'Social Agency', 'CRM Agency', 'PR Agency',
                                                       'Experimental Agency', 'Digital production Agency',
                                                       'Marketing Audits', 'Production cost (origination)',
                                                       'Production cost (adaption, transcation, repurposing)', 'SEO',
                                                       'Market Research (consumer insight)', 'Business Intelligence',
                                                       'MKT system development', 'Dealer Conference',
                                                       'Regional Meetings', 'Car cost running (depreciation)',
                                                       'System (platform) maintenance', 'DTS', 'Celebrity']:
                validation = False
    elif str(df_fcst[header_std[6]][row_idx]).upper() == 'WORKING' \
            and df_fcst[header_std[7]][row_idx] not in ['Awareness', 'Consideration', 'Opinion']:
        if df_fcst[header_std[8]][row_idx] not in ['Traditional Media', 'Digital Media', 'Social Media', 'CRM',
                                                   'Call Center', 'Event', 'Public Relationship', 'POSM',
                                                   'Sponsorship']:
            if df_fcst[header_std[9]][row_idx] not in ['TV', 'Radio', 'Newspaper & Magazine', 'Cinema', 'OOH',
                                                       'Vertical', 'Vedio', 'News & Portal', 'SEM', 'Others',
                                                       'Social Media', 'CRM Campaign', 'Consumer inbound & outbound',
                                                       'Global Event', 'Test Drive', 'Autoshow', 'Product Display',
                                                       'Launch Campaign', 'Group Buy', 'Owner experience event',
                                                       'Seasonal Campaign', 'Cyber Campaign', 'Delivery Ceremony',
                                                       'Used car campaign', 'Plant Visit', 'MY change communication',
                                                       'New product communication', 'Event communication', 'POSM',
                                                       'Customer Magazine', 'Dealer opening support', 'Sponsorship']:
                validation = False
    return validation


def validate_kpi(row_idx):
    validation = True
    if str(df_fcst[header_std[16]][row_idx]).upper() == 'Y' and str(df_fcst[header_std[11]][row_idx]) in ['0', 'nan',
                                                                                                          '/',
                                                                                                          '-'] and str(
        df_fcst[header_std[14]][row_idx]) in ['0', 'nan', '/', '-'] and str(df_fcst[header_std[13]][row_idx]) in [
        '0', 'nan', '/', '-'] and str(df_fcst[header_std[12]][row_idx]) in ['0', 'nan', '/', '-']:
        validation = False
    return validation


def is_duplicate(issue_name, row_idx):
    issue_names = [issue_name, 'Act', 'Actual']
    if df_fcst['Issue'][row_idx] not in issue_names:
        return False
    else:
        return True


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


def to_float(col_idx, row_idx):
    if 20 <= col_idx <= 32:
        if str(df_fcst[header_std[col_idx]][row_idx]).replace(' ', '') == '' or str(
                df_fcst[header_std[col_idx]][row_idx]).replace('-', '').replace(' ', '') == '':
            df_fcst[header_std[col_idx]][row_idx] = float(0.0)
        df_fcst[header_std[col_idx]][row_idx] = float(df_fcst[header_std[col_idx]][row_idx])


def to_int(col_idx, row_idx):
    if col_idx == 14:
        if not is_number(str(df_fcst[header_std[col_idx]][row_idx])):
            df_fcst[header_std[col_idx]][row_idx] = 0
        df_fcst[header_std[col_idx]][row_idx] = int(df_fcst[header_std[col_idx]][row_idx])


def to_replace(col_idx, row_idx):
    if 0 <= col_idx <= 10:
        if str(df_fcst[header_std[col_idx]][row_idx]).find('_') != -1:
            df_fcst[header_std[col_idx]][row_idx] = str(df_fcst[header_std[col_idx]][row_idx]).replace('_', ' ')


def to_replace_null(col_idx, row_idx):
    if col_idx in [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15, 16, 18, 19, 33]:
        if str(df_fcst[header_std[col_idx]][row_idx]) in ["0", "nan"]:
            df_fcst[header_std[col_idx]][row_idx] = "N.A."
    if col_idx in [11, 12, 13, 14, 17, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32]:
        if str(df_fcst[header_std[col_idx]][row_idx]) in ["0", "nan"]:
            df_fcst[header_std[col_idx]][row_idx] = int(0)


def to_split_fcstact(row_idx, f_name):
    flag = True  # True表示Act，False表示FCST
    if str(df_fcst[header_std[0]][row_idx]) != "Actual":
        flag = False
    return flag


def to_dept(dept_name, row_idx):
    flag = False
    if str(df_fcst[header_std[1]][row_idx]) == dept_name:
        flag = True
    return flag


def to_simplify(row_idx):
    df_fcst[header_std[10]][row_idx] = str(df_fcst[header_std[10]][row_idx]).replace(
        str(df_fcst[header_std[3]][row_idx]), '').lstrip()


def to_upper(row_idx, col, val):
    if str(df_fcst[header_std[col]][row_idx]).upper() == val.upper():
        df_fcst[header_std[col]][row_idx] = val


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
        column_len = data_frame[col].astype(str).str.len().mean()
        # Setting the length if the column header is larger than the max column value length
        column_len = max(column_len, len(col)) + 2
        # set the column length
        worksheet.set_column(i, i, column_len)
    if data != data_err_header and data != data_err:
        # try:
        if len(data_frame.columns) < 35:
            for row_idx in range(len(data)):
                for col_idx in range(len(header_std)):
                    if 20 <= col_idx <= 32:
                        worksheet.write(row_idx + 1, col_idx, data[row_idx][col_idx], budget_format)
        elif len(data_frame.columns) >= 35:
            for row_idx in range(len(data)):
                for col_idx in range(len(header_err)):
                    if 23 <= col_idx <= 35:
                        worksheet.write(row_idx + 1, col_idx, data[row_idx][col_idx], budget_format)
    # except:
    #     print('ok')
    writer.save()
    writer.close()


def add_error_log_and_data(f_name, row_err_data, err_data, exception, row_idx):
    # 添加error log（三列）
    row_err_data.append(f_name)
    row_err_data.append(exception)
    row_err_data.append(row_idx + 2)
    # 添加error data（header列）
    for col_idx in range(len(header_std)):
        to_replace_null(col_idx, row_idx)
        to_float(col_idx, row_idx)
        to_int(col_idx, row_idx)
        row_err_data.append(df_fcst[header_std[col_idx]][row_idx])
    err_data.append(row_err_data)


def is_null_list(checking_list1, result_list):
    if len(checking_list1) > 0:
        result_list.append(checking_list1)


def add_files(path, files):
    folder = os.listdir(path)
    for file_idx in range(len(folder)):
        if folder[file_idx].endswith('.xlsx'):
            files.append(path + folder[file_idx])


start = time.time()
# 标准数据excel header
header_std = ['Issue', 'Dept', 'Team', 'Carline', 'Lifecycle', 'Branding/NonBranding',
              'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type',
              'Activity', 'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order',
              'KPI Others', 'SMM Campaign Code (Y/N)', 'SC No.', ' SC Name',
              'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
              'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
# 异常数据excel header
header_err = ['File Name', 'Exception Type', 'Index', 'Issue', 'Dept', 'Team', 'Carline', 'Lifecycle',
              'Branding/NonBranding', 'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type', 'Activity',
              'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order', 'KPI Others', 'SMM Campaign Code (Y/N)',
              'SC No.', ' SC Name', 'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct',
              'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
# 异常header excel header
err_header_err = ['File Name', 'Exception Type']
# 读入待处理文件
path_fcst = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Data&Log/"
path_actual = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Data&Log/"
path_header = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Header Error/Files/"
file_name = []
add_files(path_fcst, file_name)
add_files(path_actual, file_name)
add_files(path_header, file_name)
data_new = []  # 存储ACT所有的标准数据
data_new_fcst = []  # 存储FCST所有的标准数据
data_err = []  # 存储ACT所有的异常数据
data_err_fcst = []  # 存储FCST所有的异常数据
data_err_header = []  # 存储所有的header异常数据
# 创建存放excel的文件夹
file_path_a = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Data&Log/"
file_path_b = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Header Error/"
file_path_c = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Header Error/Files/"
file_path_d = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Data&Log/"
mkdir(file_path_a)
mkdir(file_path_b)
mkdir(file_path_c)
mkdir(file_path_d)
# 几个生成的excel路径及文件名
today = str(datetime.date.today()).replace('-', '')
file_log = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Log.xlsx"
file_new = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Data&Log/Act_Summary_" + today + ".xlsx"
file_new_fcst = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Data&Log/FCST_Summary_" + today + ".xlsx"
file_err = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Data&Log/Act_Error_" + today + ".xlsx"
file_err_fcst = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Data&Log/FCST_Error_" + today + ".xlsx"
file_header_err = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Header Error/Header Error_" + today + ".xlsx"
# 创建几个excel的Writer引擎
writer_new = pd.ExcelWriter(path=file_new, mode='w', engine='xlsxwriter')
writer_new_fcst = pd.ExcelWriter(path=file_new_fcst, mode='w', engine='xlsxwriter')
writer_err = pd.ExcelWriter(path=file_err, mode='w', engine='xlsxwriter')
writer_err_fcst = pd.ExcelWriter(path=file_err_fcst, mode='w', engine='xlsxwriter')
writer_err_header = pd.ExcelWriter(path=file_header_err, mode='w', engine='xlsxwriter')

# 取Issue值
issue_val = pd.read_excel(file_log, dtype=str)['Issue'][0]

# 从文件列表遍历读取文件
for file_idx in range(len(file_name)):
    df_fcst = pd.read_excel(file_name[file_idx], dtype=str)
    # 判断header是否符合标准
    header_result = str(list(df_fcst.columns[0:34])).upper() == str(header_std).upper()
    if not header_result:
        row_data_err = [file_name[file_idx], 'Header Exception']
        data_err_header.append(row_data_err)
        try:
            shutil.copy(file_name[file_idx], file_path_c)
        except shutil.SameFileError:
            pass
    else:
        # 符合header标准，遍历文件的所有行，验证行数据
        for idx_row in range(len(df_fcst[header_std[0]])):
            if to_split_fcstact(idx_row, file_name[file_idx]):
                df_fcst[header_std[0]][idx_row] = 'Actual'
                row_data = []  # 临时存放符合标准的idx_row行的数据
                row_data_err = []  # 临时存放错误的idx_row行的数据
                row_data_err_null = []  # 临时存放错误的idx_row行的数据
                row_data_err_calc = []  # 临时存放错误的idx_row行的数据
                row_data_err_dupl = []  # 临时存放错误的idx_row行的数据
                row_data_err_vali = []  # 临时存放错误的idx_row行的数据
                row_data_err_ssmc = []  # 临时存放错误的idx_row行的数据
                row_data_err_kpi = []  # 临时存放错误的idx_row行的数据
                # 判断idx_row的前两列是否为空
                if not is_null(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err_null, data_err,
                                           'Issue/Dept Null Exception', idx_row)
                # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                if not calc_total_budget(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err_calc, data_err, 'TotalBudget Exception',
                                           idx_row)
                # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
                if not validate(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err_vali, data_err,
                                           'Working/NonWorking Exception', idx_row)
                if not is_ssm(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err_ssmc, data_err, 'SMM Exception', idx_row)
                if not validate_kpi(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err_kpi, data_err, 'KPI Exception', idx_row)
                # 符合所有标准的数据格式化后写入Summary文件
                if is_null(idx_row) and calc_total_budget(idx_row) and validate(idx_row) and is_ssm(idx_row) and validate_kpi(idx_row):
                    # df_fcst[header_std[0]][idx_row] = 'Actual'
                    for idx_col in range(len(header_std)):
                        to_replace(idx_col, idx_row)
                        to_simplify(idx_row)
                        to_replace_null(idx_col, idx_row)
                        to_upper(idx_row, 5, 'Branding')
                        to_upper(idx_row, 5, 'NonBranding')
                        to_upper(idx_row, 6, 'Working')
                        to_upper(idx_row, 6, 'NonWorking')
                        to_float(idx_col, idx_row)
                        to_int(idx_col, idx_row)
                        row_data.append(df_fcst[header_std[idx_col]][idx_row])
                    data_new.append(row_data)
            else:
                row_data = []  # 临时存放符合标准的idx_row行的数据
                row_data_err = []  # 临时存放错误的idx_row行的数据
                row_data_err_null_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_calc_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_dupl_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_vali_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_ssmc_fcst = []  # 临时存放错误的idx_row行的数据
                row_data_err_kpi_fcst = []  # 临时存放错误的idx_row行的数据
                # 判断idx_row的前两列是否为空
                if not is_null(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err_null_fcst, data_err_fcst,
                                           'Issue/Dept Null Exception', idx_row)
                # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                if not calc_total_budget(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err_calc_fcst, data_err_fcst,
                                           'TotalBudget Exception', idx_row)
                # 判断idx_row的Issue列是否含有多余值（Issue列值除【文件名内包含的issue值和Act】，其余都视为多余值）
                if not is_duplicate(issue_val, idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err_dupl_fcst, data_err_fcst,
                                           'Issue Exception', idx_row)
                if not validate(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err_vali_fcst, data_err_fcst,
                                           'Working/NonWorking Exception', idx_row)
                if not is_ssm(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err_ssmc_fcst, data_err_fcst, 'SMM Exception',
                                           idx_row)
                if not validate_kpi(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err_kpi_fcst, data_err_fcst, 'KPI Exception',
                                           idx_row)
                # 符合所有标准的数据格式化后写入Summary文件
                if is_null(idx_row) and calc_total_budget(idx_row) and is_duplicate(issue_val,
                                                                                    idx_row) and validate(
                    idx_row) and is_ssm(idx_row) and validate_kpi(idx_row):
                    for idx_col in range(len(header_std)):
                        to_replace(idx_col, idx_row)
                        to_simplify(idx_row)
                        to_replace_null(idx_col, idx_row)
                        to_upper(idx_row, 5, 'Branding')
                        to_upper(idx_row, 5, 'NonBranding')
                        to_upper(idx_row, 6, 'Working')
                        to_upper(idx_row, 6, 'NonWorking')
                        to_float(idx_col, idx_row)
                        to_int(idx_col, idx_row)
                        row_data.append(df_fcst[header_std[idx_col]][idx_row])
                    data_new_fcst.append(row_data)
# 将异常header写入Error Header excel
frame_header_err = pd.DataFrame(data_err_header, columns=err_header_err)
frame_header_err.to_excel(writer_err_header, index=False, header=True)
to_format(writer_err_header, frame_header_err, data_err_header, [0, 1], [], [])
# 将Actual异常数据写入Error excel
frame_err = pd.DataFrame(data_err, columns=header_err)
frame_err.to_excel(writer_err, index=False, header=True)
to_format(writer_err, frame_err, data_err, [0, 1, 2], [18, 19, 22, 35], [20, 21])
# 将FCST异常数据写入Error excel
frame_err_fcst = pd.DataFrame(data_err_fcst, columns=header_err)
frame_err_fcst.to_excel(writer_err_fcst, index=False, header=True)
to_format(writer_err_fcst, frame_err_fcst, data_err_fcst, [0, 1, 2], [18, 19, 22, 35], [20, 21])
# 将Actual标准数据写入Error excel
frame_new = pd.DataFrame(data_new, columns=header_std)
frame_new.to_excel(writer_new, index=False, header=True)
to_format(writer_new, frame_new, data_new, [], [15, 16, 19, 32], [17, 18])
# 将FCST标准数据写入Error excel
frame_new_fcst = pd.DataFrame(data_new_fcst, columns=header_std)
frame_new_fcst.to_excel(writer_new_fcst, index=False, header=True)
to_format(writer_new_fcst, frame_new_fcst, data_new_fcst, [], [15, 16, 19, 32], [17, 18])


def to_float2(col_idx, row_idx):
    if 20 <= col_idx <= 32:
        df_summary[header_std[col_idx]][row_idx] = float(df_summary[header_std[col_idx]][row_idx])


def to_format2(writer, data_frame, data, red_col, gray_col):
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
    budget_format = workbook.add_format({'num_format': '#,##0.00'})
    # Write the column headers with the defined format.
    for col_num, value in enumerate(data_frame.columns.values):
        if col_num in red_col:
            worksheet.write(0, col_num, value, header_format_pqt)
        elif col_num in gray_col:
            worksheet.write(0, col_num, value, header_format_rs)
        else:
            worksheet.write(0, col_num, value, header_format)
    # Iterate through each column and set the width == the max length in that column.
    # A padding length of 2 is also added.
    for i, col in enumerate(data_frame.columns):
        # find length of column i
        column_len = data_frame[col].astype(str).str.len().mean()
        # Setting the length if the column header is larger than the max column value length
        column_len = max(column_len, len(col)) + 2
        # set the column length
        worksheet.set_column(i, i, column_len)
    for row_idx in range(len(data)):
        for col_idx in range(len(header_std)):
            if 20 <= col_idx <= 32:
                worksheet.write(row_idx + 1, col_idx, data[row_idx][col_idx], budget_format)
    writer.save()
    writer.close()


path1 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/"
mkdir(path1)
# 读入待处理文件
file_summary = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Data&Log/FCST_Summary_" + today + ".xlsx"
df_summary = pd.read_excel(file_summary, dtype=str)

file_MKT = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/FCST_MKT_" + today + ".xlsx"
file_MKT_Launch = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/FCST_MKT Launch_" + today + ".xlsx"
file_APAC_CC = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/FCST_APAC CC_" + today + ".xlsx"
file_Customer_Service = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/FCST_Customer Service_" + today + ".xlsx"
file_DTS = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/FCST_DTS_" + today + ".xlsx"
file_MI = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/FCST_MI_" + today + ".xlsx"
file_Sales_MKT1 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/FCST_Sales MKT_DMKT_" + today + ".xlsx"
file_Sales_MKT2 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/FCST_Sales MKT_HK_" + today + ".xlsx"
file_Sales_MKT3 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/FCST_Sales MKT_Internal Fleet Car_" + today + ".xlsx"
file_Sales_MKT4 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/FCST_Sales MKT_National Sales_" + today + ".xlsx"
file_Sales_MKT5 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/FCST/Dept-Team/FCST_Sales MKT_NBD Team_" + today + ".xlsx"

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

# 创建几个excel的Writer引擎
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

# 符合header标准，遍历文件的所有行，验证行数据
for idx_row in range(len(df_summary[header_std[0]])):
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
    if str(df_summary[header_std[1]][idx_row]) == 'MKT':
        for idx_col in range(len(header_std)):
            to_float2(idx_col, idx_row)
            row_data_MKT.append(df_summary[header_std[idx_col]][idx_row])
        data_MKT.append(row_data_MKT)
    elif str(df_summary[header_std[1]][idx_row]) == 'MKT Launch':
        for idx_col in range(len(header_std)):
            to_float2(idx_col, idx_row)
            row_data_MKT_Launch.append(df_summary[header_std[idx_col]][idx_row])
        data_MKT_Launch.append(row_data_MKT_Launch)
    elif str(df_summary[header_std[1]][idx_row]) == 'APAC CC':
        for idx_col in range(len(header_std)):
            to_float2(idx_col, idx_row)
            row_data_APAC_CC.append(df_summary[header_std[idx_col]][idx_row])
        data_APAC_CC.append(row_data_APAC_CC)
    elif str(df_summary[header_std[1]][idx_row]) == 'Customer Service':
        for idx_col in range(len(header_std)):
            to_float2(idx_col, idx_row)
            row_data_Customer_Service.append(df_summary[header_std[idx_col]][idx_row])
        data_Customer_Service.append(row_data_Customer_Service)
    elif str(df_summary[header_std[1]][idx_row]) == 'DTS':
        for idx_col in range(len(header_std)):
            to_float2(idx_col, idx_row)
            row_data_DTS.append(df_summary[header_std[idx_col]][idx_row])
        data_DTS.append(row_data_DTS)
    elif str(df_summary[header_std[1]][idx_row]) == 'MI':
        for idx_col in range(len(header_std)):
            to_float2(idx_col, idx_row)
            row_data_MI.append(df_summary[header_std[idx_col]][idx_row])
        data_MI.append(row_data_MI)
    elif str(df_summary[header_std[1]][idx_row]) in ['Sales MKT', 'Sales_MKT'] and str(
            df_summary[header_std[2]][idx_row]) in [
        'DMKT Central Team', 'Exhibition', 'East Region Team', 'North Region Team', 'South Region Team',
        'West Region Team', 'ZheJiang Region Team']:
        for idx_col in range(len(header_std)):
            to_float2(idx_col, idx_row)
            row_data_Sales_MKT1.append(df_summary[header_std[idx_col]][idx_row])
        data_Sales_MKT1.append(row_data_Sales_MKT1)
    elif str(df_summary[header_std[1]][idx_row]) in ['Sales MKT', 'Sales_MKT'] and str(
            df_summary[header_std[2]][idx_row]) in ['HK']:
        for idx_col in range(len(header_std)):
            to_float2(idx_col, idx_row)
            row_data_Sales_MKT2.append(df_summary[header_std[idx_col]][idx_row])
        data_Sales_MKT2.append(row_data_Sales_MKT2)
    elif str(df_summary[header_std[1]][idx_row]) in ['Sales MKT', 'Sales_MKT'] and str(
            df_summary[header_std[2]][idx_row]) in [
        'Internal Fleet Car']:
        for idx_col in range(len(header_std)):
            to_float2(idx_col, idx_row)
            row_data_Sales_MKT3.append(df_summary[header_std[idx_col]][idx_row])
        data_Sales_MKT3.append(row_data_Sales_MKT3)
    elif str(df_summary[header_std[1]][idx_row]) in ['Sales MKT', 'Sales_MKT'] and str(
            df_summary[header_std[2]][idx_row]) in [
        'National Sales']:
        for idx_col in range(len(header_std)):
            to_float2(idx_col, idx_row)
            row_data_Sales_MKT4.append(df_summary[header_std[idx_col]][idx_row])
        data_Sales_MKT4.append(row_data_Sales_MKT4)
    elif str(df_summary[header_std[1]][idx_row]) in ['Sales MKT', 'Sales_MKT'] and str(
            df_summary[header_std[2]][idx_row]) in [
        'NBD Team']:
        for idx_col in range(len(header_std)):
            to_float2(idx_col, idx_row)
            row_data_Sales_MKT5.append(df_summary[header_std[idx_col]][idx_row])
        data_Sales_MKT5.append(row_data_Sales_MKT5)

# 将标准数据写入Summary excel
frame_new1 = pd.DataFrame(data_MKT, columns=header_std)
frame_new1.to_excel(writer_MKT, index=False, header=True)
to_format2(writer_MKT, frame_new1, data_MKT, [15, 16, 19, 32], [17, 18])

frame_new2 = pd.DataFrame(data_MKT_Launch, columns=header_std)
frame_new2.to_excel(writer_MKT_Launch, index=False, header=True)
to_format2(writer_MKT_Launch, frame_new2, data_MKT_Launch, [15, 16, 19, 32], [17, 18])

frame_new3 = pd.DataFrame(data_APAC_CC, columns=header_std)
frame_new3.to_excel(writer_APAC_CC, index=False, header=True)
to_format2(writer_APAC_CC, frame_new3, data_APAC_CC, [15, 16, 19, 32], [17, 18])

frame_new4 = pd.DataFrame(data_Customer_Service, columns=header_std)
frame_new4.to_excel(writer_Customer_Service, index=False, header=True)
to_format2(writer_Customer_Service, frame_new4, data_Customer_Service, [15, 16, 19, 32], [17, 18])

frame_new5 = pd.DataFrame(data_DTS, columns=header_std)
frame_new5.to_excel(writer_DTS, index=False, header=True)
to_format2(writer_DTS, frame_new5, data_DTS, [15, 16, 19, 32], [17, 18])

frame_new6 = pd.DataFrame(data_MI, columns=header_std)
frame_new6.to_excel(writer_MI, index=False, header=True)
to_format2(writer_MI, frame_new6, data_MI, [15, 16, 19, 32], [17, 18])

frame_new7 = pd.DataFrame(data_Sales_MKT1, columns=header_std)
frame_new7.to_excel(writer_Sales_MKT1, index=False, header=True)
to_format2(writer_Sales_MKT1, frame_new7, data_Sales_MKT1, [15, 16, 19, 32], [17, 18])

frame_new8 = pd.DataFrame(data_Sales_MKT2, columns=header_std)
frame_new8.to_excel(writer_Sales_MKT2, index=False, header=True)
to_format2(writer_Sales_MKT2, frame_new8, data_Sales_MKT2, [15, 16, 19, 32], [17, 18])

frame_new9 = pd.DataFrame(data_Sales_MKT3, columns=header_std)
frame_new9.to_excel(writer_Sales_MKT3, index=False, header=True)
to_format2(writer_Sales_MKT3, frame_new9, data_Sales_MKT3, [15, 16, 19, 32], [17, 18])

frame_new10 = pd.DataFrame(data_Sales_MKT4, columns=header_std)
frame_new10.to_excel(writer_Sales_MKT4, index=False, header=True)
to_format2(writer_Sales_MKT4, frame_new10, data_Sales_MKT4, [15, 16, 19, 32], [17, 18])

frame_new11 = pd.DataFrame(data_Sales_MKT5, columns=header_std)
frame_new11.to_excel(writer_Sales_MKT5, index=False, header=True)
to_format2(writer_Sales_MKT5, frame_new11, data_Sales_MKT5, [15, 16, 19, 32], [17, 18])


def to_float3(col_idx, row_idx):
    if 20 <= col_idx <= 32:
        df_summary2[header_std[col_idx]][row_idx] = float(df_summary2[header_std[col_idx]][row_idx])


def to_format2(writer, data_frame, data, red_col, gray_col):
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
    budget_format = workbook.add_format({'num_format': '#,##0.00'})
    # Write the column headers with the defined format.
    for col_num, value in enumerate(data_frame.columns.values):
        if col_num in red_col:
            worksheet.write(0, col_num, value, header_format_pqt)
        elif col_num in gray_col:
            worksheet.write(0, col_num, value, header_format_rs)
        else:
            worksheet.write(0, col_num, value, header_format)
    # Iterate through each column and set the width == the max length in that column.
    # A padding length of 2 is also added.
    for i, col in enumerate(data_frame.columns):
        # find length of column i
        column_len = data_frame[col].astype(str).str.len().mean()
        # Setting the length if the column header is larger than the max column value length
        column_len = max(column_len, len(col)) + 2
        # set the column length
        worksheet.set_column(i, i, column_len)
    for row_idx in range(len(data)):
        for col_idx in range(len(header_std)):
            if 20 <= col_idx <= 32:
                worksheet.write(row_idx + 1, col_idx, data[row_idx][col_idx], budget_format)
    writer.save()
    writer.close()


path2 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/"
mkdir(path2)
# 读入待处理文件
file_summary2 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Data&Log/Act_Summary_" + today + ".xlsx"
df_summary2 = pd.read_excel(file_summary2, dtype=str)

file_MKT2 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/Act_MKT_" + today + ".xlsx"
file_MKT_Launch2 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/Act_MKT Launch_" + today + ".xlsx"
file_APAC_CC2 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/Act_APAC CC_" + today + ".xlsx"
file_Customer_Service2 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/Act_Customer Service_" + today + ".xlsx"
file_DTS2 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/Act_DTS_" + today + ".xlsx"
file_MI2 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/Act_MI_" + today + ".xlsx"
file_Sales_MKT12 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/Act_Sales MKT_DMKT_" + today + ".xlsx"
file_Sales_MKT22 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/Act_Sales MKT_HK_" + today + ".xlsx"
file_Sales_MKT32 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/Act_Sales MKT_Internal Fleet Car_" + today + ".xlsx"
file_Sales_MKT42 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/Act_Sales MKT_National Sales_" + today + ".xlsx"
file_Sales_MKT52 = "W:/Finance Volvo/Confidential/Albert/6 QLIKVIEW/QlikView_sources_restricted/ASP/ASP Testing/Actual/Dept-Team/Act_Sales MKT_NBD Team_" + today + ".xlsx"

data_MKT2 = []  # 存储所有的标准数据
data_MKT_Launch2 = []
data_APAC_CC2 = []
data_Customer_Service2 = []
data_DTS2 = []
data_MI2 = []
data_Sales_MKT12 = []
data_Sales_MKT22 = []
data_Sales_MKT32 = []
data_Sales_MKT42 = []
data_Sales_MKT52 = []

# 创建几个excel的Writer引擎
writer_MKT2 = pd.ExcelWriter(path=file_MKT2, mode='w', engine='xlsxwriter')
writer_MKT_Launch2 = pd.ExcelWriter(path=file_MKT_Launch2, mode='w', engine='xlsxwriter')
writer_APAC_CC2 = pd.ExcelWriter(path=file_APAC_CC2, mode='w', engine='xlsxwriter')
writer_Customer_Service2 = pd.ExcelWriter(path=file_Customer_Service2, mode='w', engine='xlsxwriter')
writer_DTS2 = pd.ExcelWriter(path=file_DTS2, mode='w', engine='xlsxwriter')
writer_MI2 = pd.ExcelWriter(path=file_MI2, mode='w', engine='xlsxwriter')
writer_Sales_MKT12 = pd.ExcelWriter(path=file_Sales_MKT12, mode='w', engine='xlsxwriter')
writer_Sales_MKT22 = pd.ExcelWriter(path=file_Sales_MKT22, mode='w', engine='xlsxwriter')
writer_Sales_MKT32 = pd.ExcelWriter(path=file_Sales_MKT32, mode='w', engine='xlsxwriter')
writer_Sales_MKT42 = pd.ExcelWriter(path=file_Sales_MKT42, mode='w', engine='xlsxwriter')
writer_Sales_MKT52 = pd.ExcelWriter(path=file_Sales_MKT52, mode='w', engine='xlsxwriter')

# 符合header标准，遍历文件的所有行，验证行数据
for idx_row in range(len(df_summary2[header_std[0]])):
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
    if str(df_summary2[header_std[1]][idx_row]) == 'MKT':
        for idx_col in range(len(header_std)):
            to_float3(idx_col, idx_row)
            row_data_MKT.append(df_summary2[header_std[idx_col]][idx_row])
        data_MKT2.append(row_data_MKT)
    elif str(df_summary2[header_std[1]][idx_row]) == 'MKT Launch':
        for idx_col in range(len(header_std)):
            to_float3(idx_col, idx_row)
            row_data_MKT_Launch.append(df_summary2[header_std[idx_col]][idx_row])
        data_MKT_Launch2.append(row_data_MKT_Launch)
    elif str(df_summary2[header_std[1]][idx_row]) == 'APAC CC':
        for idx_col in range(len(header_std)):
            to_float3(idx_col, idx_row)
            row_data_APAC_CC.append(df_summary2[header_std[idx_col]][idx_row])
        data_APAC_CC2.append(row_data_APAC_CC)
    elif str(df_summary2[header_std[1]][idx_row]) == 'Customer Service':
        for idx_col in range(len(header_std)):
            to_float3(idx_col, idx_row)
            row_data_Customer_Service.append(df_summary2[header_std[idx_col]][idx_row])
        data_Customer_Service2.append(row_data_Customer_Service)
    elif str(df_summary2[header_std[1]][idx_row]) == 'DTS':
        for idx_col in range(len(header_std)):
            to_float3(idx_col, idx_row)
            row_data_DTS.append(df_summary2[header_std[idx_col]][idx_row])
        data_DTS2.append(row_data_DTS)
    elif str(df_summary2[header_std[1]][idx_row]) == 'MI':
        for idx_col in range(len(header_std)):
            to_float3(idx_col, idx_row)
            row_data_MI.append(df_summary2[header_std[idx_col]][idx_row])
        data_MI2.append(row_data_MI)
    elif str(df_summary2[header_std[1]][idx_row]) in ['Sales MKT', 'Sales_MKT'] and str(
            df_summary2[header_std[2]][idx_row]) in [
        'DMKT Central Team', 'Exhibition', 'East Region Team', 'North Region Team', 'South Region Team',
        'West Region Team', 'ZheJiang Region Team']:
        for idx_col in range(len(header_std)):
            to_float3(idx_col, idx_row)
            row_data_Sales_MKT1.append(df_summary2[header_std[idx_col]][idx_row])
        data_Sales_MKT12.append(row_data_Sales_MKT1)
    elif str(df_summary2[header_std[1]][idx_row]) in ['Sales MKT', 'Sales_MKT'] and str(
            df_summary2[header_std[2]][idx_row]) in ['HK']:
        for idx_col in range(len(header_std)):
            to_float3(idx_col, idx_row)
            row_data_Sales_MKT2.append(df_summary2[header_std[idx_col]][idx_row])
        data_Sales_MKT22.append(row_data_Sales_MKT2)
    elif str(df_summary2[header_std[1]][idx_row]) in ['Sales MKT', 'Sales_MKT'] and str(
            df_summary2[header_std[2]][idx_row]) in [
        'Internal Fleet Car']:
        for idx_col in range(len(header_std)):
            to_float3(idx_col, idx_row)
            row_data_Sales_MKT3.append(df_summary2[header_std[idx_col]][idx_row])
        data_Sales_MKT32.append(row_data_Sales_MKT3)
    elif str(df_summary2[header_std[1]][idx_row]) in ['Sales MKT', 'Sales_MKT'] and str(
            df_summary2[header_std[2]][idx_row]) in [
        'National Sales']:
        for idx_col in range(len(header_std)):
            to_float3(idx_col, idx_row)
            row_data_Sales_MKT4.append(df_summary2[header_std[idx_col]][idx_row])
        data_Sales_MKT42.append(row_data_Sales_MKT4)
    elif str(df_summary2[header_std[1]][idx_row]) in ['Sales MKT', 'Sales_MKT'] and str(
            df_summary2[header_std[2]][idx_row]) in [
        'NBD Team']:
        for idx_col in range(len(header_std)):
            to_float3(idx_col, idx_row)
            row_data_Sales_MKT5.append(df_summary2[header_std[idx_col]][idx_row])
        data_Sales_MKT52.append(row_data_Sales_MKT5)

# 将标准数据写入Summary excel
frame_new12 = pd.DataFrame(data_MKT2, columns=header_std)
frame_new12.to_excel(writer_MKT2, index=False, header=True)
to_format2(writer_MKT2, frame_new12, data_MKT2, [15, 16, 19, 32], [17, 18])

frame_new22 = pd.DataFrame(data_MKT_Launch2, columns=header_std)
frame_new22.to_excel(writer_MKT_Launch2, index=False, header=True)
to_format2(writer_MKT_Launch2, frame_new22, data_MKT_Launch2, [15, 16, 19, 32], [17, 18])

frame_new32 = pd.DataFrame(data_APAC_CC2, columns=header_std)
frame_new32.to_excel(writer_APAC_CC2, index=False, header=True)
to_format2(writer_APAC_CC2, frame_new32, data_APAC_CC2, [15, 16, 19, 32], [17, 18])

frame_new42 = pd.DataFrame(data_Customer_Service2, columns=header_std)
frame_new42.to_excel(writer_Customer_Service2, index=False, header=True)
to_format2(writer_Customer_Service2, frame_new42, data_Customer_Service2, [15, 16, 19, 32], [17, 18])

frame_new52 = pd.DataFrame(data_DTS2, columns=header_std)
frame_new52.to_excel(writer_DTS2, index=False, header=True)
to_format2(writer_DTS2, frame_new52, data_DTS2, [15, 16, 19, 32], [17, 18])

frame_new62 = pd.DataFrame(data_MI2, columns=header_std)
frame_new62.to_excel(writer_MI2, index=False, header=True)
to_format2(writer_MI2, frame_new62, data_MI2, [15, 16, 19, 32], [17, 18])

frame_new72 = pd.DataFrame(data_Sales_MKT12, columns=header_std)
frame_new72.to_excel(writer_Sales_MKT12, index=False, header=True)
to_format2(writer_Sales_MKT12, frame_new72, data_Sales_MKT12, [15, 16, 19, 32], [17, 18])

frame_new82 = pd.DataFrame(data_Sales_MKT22, columns=header_std)
frame_new82.to_excel(writer_Sales_MKT22, index=False, header=True)
to_format2(writer_Sales_MKT22, frame_new82, data_Sales_MKT22, [15, 16, 19, 32], [17, 18])

frame_new92 = pd.DataFrame(data_Sales_MKT32, columns=header_std)
frame_new92.to_excel(writer_Sales_MKT32, index=False, header=True)
to_format2(writer_Sales_MKT32, frame_new92, data_Sales_MKT32, [15, 16, 19, 32], [17, 18])

frame_new102 = pd.DataFrame(data_Sales_MKT42, columns=header_std)
frame_new102.to_excel(writer_Sales_MKT42, index=False, header=True)
to_format2(writer_Sales_MKT42, frame_new102, data_Sales_MKT42, [15, 16, 19, 32], [17, 18])

frame_new112 = pd.DataFrame(data_Sales_MKT52, columns=header_std)
frame_new112.to_excel(writer_Sales_MKT52, index=False, header=True)
to_format2(writer_Sales_MKT52, frame_new112, data_Sales_MKT52, [15, 16, 19, 32], [17, 18])
# 计算程序耗时
cost = time.time() - start
print(cost)
# 生成log
try:
    log_time = []
    log_col = ['LastUpdatedTime', 'Issue']
    log_time.append([str(datetime.datetime.now()),str(frame_new_fcst['Issue'][0])])
    writer_log = pd.ExcelWriter(path=file_log, mode='w', engine='xlsxwriter')
    frame_log = pd.DataFrame(log_time, columns=log_col)
    frame_log.to_excel(writer_log, index=False, header=True)
    writer_log.save()
    writer_log.close()
except:
    pass
print("Process finished with exit code 0")