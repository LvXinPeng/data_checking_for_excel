import time
import pandas as pd
import os
import shutil


def mkdir(file_path):
    if not os.path.exists(file_path):
        os.makedirs(file_path)


def is_null(row_idx):
    is_na = True
    if str(df_fcst[header_std[0]][row_idx]) == "nan" or str(df_fcst[header_std[1]][row_idx]) == "nan":
        is_na = False
    return is_na


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
    if str(df_fcst[header_std[6]][row_idx]).upper() == 'NONWORKING' and str(df_fcst[header_std[7]][row_idx]) not in \
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


def is_duplicate(f_name, row_idx):
    # eg: _issue.xlsx
    issue_name = f_name.split('.')[0].replace(path, '')[0:8]
    issue_names = [issue_name, 'Act', 'Actual']
    if df_fcst['Issue'][row_idx] not in issue_names:
        return False
    else:
        return True


def to_float(col_idx, row_idx):
    if 20 <= col_idx <= 32:
        if str(df_fcst[header_std[col_idx]][row_idx]).replace(' ', '') == '' or str(
                df_fcst[header_std[col_idx]][row_idx]).replace('-', '').replace(' ', '') == '':
            df_fcst[header_std[col_idx]][row_idx] = float(0.0)
        df_fcst[header_std[col_idx]][row_idx] = float(df_fcst[header_std[col_idx]][row_idx])


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
    issue_name = f_name.split('.')[0].replace(path, '')[0:8]
    if str(df_fcst[header_std[0]][row_idx]) == issue_name and str(df_fcst[header_std[17]][row_idx]) in ['nan', '0', '',
                                                                                                        ' ']:
        df_fcst[header_std[17]][row_idx] = 0
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
        row_err_data.append(df_fcst[header_std[col_idx]][row_idx])
    err_data.append(row_err_data)


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
path = "D:/ASP - Erin/Raw Data/"
file_name = os.listdir('D:/ASP - Erin/Raw Data/')
for file_idx in range(len(file_name)):
    file_name[file_idx] = path + file_name[file_idx]
# file_name = input()
data_new = []  # 存储ACT所有的标准数据
data_new_fcst = []  # 存储FCST所有的标准数据
data_err = []  # 存储ACT所有的异常数据
data_err_fcst = []  # 存储FCST所有的异常数据
data_err_header = []  # 存储所有的header异常数据
# 创建存放excel的文件夹
file_path_a = "D:/ASP - Erin/FCST/Data&Log/"
file_path_b = "D:/ASP - Erin/Header Error/"
file_path_c = "D:/ASP - Erin/Header Error/Files/"
file_path_d = "D:/ASP - Erin/Actual/Data&Log/"
mkdir(file_path_a)
mkdir(file_path_b)
mkdir(file_path_c)
mkdir(file_path_d)
# 几个生成的excel路径及文件名
file_new = "D:/ASP - Erin/Actual/Data&Log/Summary.xlsx"
file_new_fcst = "D:/ASP - Erin/FCST/Data&Log/Summary.xlsx"
file_err = "D:/ASP - Erin/Actual/Data&Log/Error.xlsx"
file_err_fcst = "D:/ASP - Erin/FCST/Data&Log/Error.xlsx"
file_header_err = "D:/ASP - Erin/Header Error/Header Error.xlsx"
# 创建几个excel的Writer引擎
writer_new = pd.ExcelWriter(path=file_new, mode='w', engine='xlsxwriter')
writer_new_fcst = pd.ExcelWriter(path=file_new_fcst, mode='w', engine='xlsxwriter')
writer_err = pd.ExcelWriter(path=file_err, mode='w', engine='xlsxwriter')
writer_err_fcst = pd.ExcelWriter(path=file_err_fcst, mode='w', engine='xlsxwriter')
writer_err_header = pd.ExcelWriter(path=file_header_err, mode='w', engine='xlsxwriter')

# 从文件列表遍历读取文件
for file_idx in range(len(file_name)):
    df_fcst = pd.read_excel(file_name[file_idx], dtype=str)
    # 判断header是否符合标准
    header_result = df_fcst.columns[0:34] == header_std
    if not header_result.all():
        row_data_err = [file_name[file_idx], 'Header Exception']
        data_err_header.append(row_data_err)
        shutil.copy(file_name[file_idx], file_path_c)
    else:
        # 符合header标准，遍历文件的所有行，验证行数据
        for idx_row in range(len(df_fcst[header_std[0]])):
            if to_split_fcstact(idx_row, file_name[file_idx]):
                row_data = []  # 临时存放符合标准的idx_row行的数据
                row_data_err = []  # 临时存放错误的idx_row行的数据
                # 判断idx_row的前两列是否为空
                if not is_null(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err, data_err, 'Issue/Dept Null Exception',
                                           idx_row)
                # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                elif not calc_total_budget(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err, data_err, 'TotalBudget Exception',
                                           idx_row)
                # 判断idx_row的Issue列是否含有多余值（Issue列值除【文件名内包含的issue值和Act】，其余都视为多余值）
                elif not is_duplicate(file_name[file_idx], idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err, data_err, 'Issue Exception', idx_row)
                # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
                elif not validate(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err, data_err, 'Working/NonWorking Exception',
                                           idx_row)
                # 符合所有标准的数据格式化后写入Summary文件
                else:
                    df_fcst[header_std[0]][idx_row] = 'Actual'
                    for idx_col in range(len(header_std)):
                        to_replace(idx_col, idx_row)
                        to_simplify(idx_row)
                        to_replace_null(idx_col, idx_row)
                        to_upper(idx_row, 5, 'Branding')
                        to_upper(idx_row, 5, 'NonBranding')
                        to_upper(idx_row, 6, 'Working')
                        to_upper(idx_row, 6, 'NonWorking')
                        to_float(idx_col, idx_row)
                        row_data.append(df_fcst[header_std[idx_col]][idx_row])
                    data_new.append(row_data)
            else:
                row_data = []  # 临时存放符合标准的idx_row行的数据
                row_data_err = []  # 临时存放错误的idx_row行的数据
                # 判断idx_row的前两列是否为空
                if not is_null(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err, data_err_fcst,
                                           'Issue/Dept Null Exception',
                                           idx_row)
                # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                elif not calc_total_budget(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err, data_err_fcst, 'TotalBudget Exception',
                                           idx_row)
                # 判断idx_row的Issue列是否含有多余值（Issue列值除【文件名内包含的issue值和Act】，其余都视为多余值）
                elif not is_duplicate(file_name[file_idx], idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err, data_err_fcst, 'Issue Exception', idx_row)
                # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
                elif not validate(idx_row):
                    add_error_log_and_data(file_name[file_idx], row_data_err, data_err_fcst,
                                           'Working/NonWorking Exception',
                                           idx_row)
                # 符合所有标准的数据格式化后写入Summary文件
                else:
                    for idx_col in range(len(header_std)):
                        to_replace(idx_col, idx_row)
                        to_simplify(idx_row)
                        to_replace_null(idx_col, idx_row)
                        to_upper(idx_row, 5, 'Branding')
                        to_upper(idx_row, 5, 'NonBranding')
                        to_upper(idx_row, 6, 'Working')
                        to_upper(idx_row, 6, 'NonWorking')
                        to_float(idx_col, idx_row)
                        row_data.append(df_fcst[header_std[idx_col]][idx_row])
                    data_new_fcst.append(row_data)


# 将异常header写入Error Header excel
frame_header_err = pd.DataFrame(data_err_header, columns=err_header_err)
frame_header_err.to_excel(writer_err_header, index=False, header=True)
to_format(writer_err_header, frame_header_err, data_err_header, [0, 1], [], [])
# 将异常数据写入Error excel
frame_err = pd.DataFrame(data_err, columns=header_err)
frame_err.to_excel(writer_err, index=False, header=True)
to_format(writer_err, frame_err, data_err, [0, 1, 2], [18, 19, 22, 35], [20, 21])

frame_err_fcst = pd.DataFrame(data_err_fcst, columns=header_err)
frame_err_fcst.to_excel(writer_err_fcst, index=False, header=True)
to_format(writer_err_fcst, frame_err_fcst, data_err_fcst, [0, 1, 2], [18, 19, 22, 35], [20, 21])

frame_new = pd.DataFrame(data_new, columns=header_std)
frame_new.to_excel(writer_new, index=False, header=True)
to_format(writer_new, frame_new, data_new, [], [15, 16, 19, 32], [17, 18])

frame_new_fcst = pd.DataFrame(data_new_fcst, columns=header_std)
frame_new_fcst.to_excel(writer_new_fcst, index=False, header=True)
to_format(writer_new_fcst, frame_new_fcst, data_new_fcst, [], [15, 16, 19, 32], [17, 18])
# 计算程序耗时
cost = time.time() - start
print(cost)