import time
import pandas as pd
import os


def is_null(row_idx):
    is_na = True
    if str(df_fcst[header_std[0]][row_idx]) == "nan" or str(df_fcst[header_std[1]][row_idx]) == "nan":
        is_na = False
    return is_na


def calc_total_budget(row_idx):
    temp = []
    for col_idx in range(12):
        if str(df_fcst[header_std[col_idx + 20]][row_idx]).replace(' ', '') == '':
            temp.append(0)
        elif str(df_fcst[header_std[col_idx + 20]][row_idx]).replace('-', '').replace(' ', '') == '':
            temp.append(0)
        else:
            temp.append(df_fcst[header_std[col_idx + 20]][row_idx])
    total_budget = sum(list(map(float, temp)))
    is_eq = total_budget == float(df_fcst[header_std[32]][row_idx])
    return is_eq


def validate(row_idx):
    validation = True
    if str(df_fcst[header_std[6]][row_idx]).upper() == 'NONWORKING' and df_fcst[header_std[7]][row_idx] != "0":
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
    elif str(df_fcst[header_std[6]][row_idx]).upper() == 'WORKING' and df_fcst[header_std[7]][row_idx] not in [
        'Awareness', 'Consideration', 'Opinion']:
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
    issue_name = f_name.split('.')[0][-9:]
    issue_names = [issue_name, 'Act']
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


def to_upper(row_idx, col, val):
    if str(df_fcst[header_std[col]][row_idx]).upper() == val.upper():
        df_fcst[header_std[col]][row_idx] = val


def to_format(writer, data_frame, err_col, red_col, gray_col):
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
    writer.save()
    writer.close()


def add_error_log_and_data(f_name, row_err_data, err_data, exception, row_idx):
    # 添加error log（三列）
    row_err_data.append(f_name)
    row_err_data.append(exception)
    row_err_data.append(row_idx + 2)
    # 添加error data（header列）
    for idx_col in range(len(header_std)):
        to_float(idx_col, row_idx)
        row_err_data.append(df_fcst[header_std[idx_col]][row_idx])
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
path = "D:/ASP - Erin/Raw Data/checking list.xlsx"
df_files = pd.read_excel(path, dtype=str)
file_name = df_files['File Path']
# file_name = input()
data_new = []  # 存储所有的标准数据
data_err = []  # 存储所有的异常数据
data_err_header = []  # 存储所有的header异常数据
# 创建存放excel的文件夹
file_path_a = "D:/ASP - Erin/Raw Data/Data&Log/"
file_path_b = "D:/ASP - Erin/Raw Data/Error/"
if not os.path.exists(file_path_a):
    os.makedirs(file_path_a)
if not os.path.exists(file_path_b):
    os.makedirs(file_path_b)
# 几个生成的excel路径及文件名
file_new = "D:/ASP - Erin/Raw Data/Data&Log/Summary.xlsx"
file_err = "D:/ASP - Erin/Raw Data/Data&Log/Error.xlsx"
file_header_err = "D:/ASP - Erin/Raw Data/Error/Error_header.xlsx"
# 创建几个excel的Writer引擎
writer_new = pd.ExcelWriter(path=file_new, mode='w', engine='xlsxwriter')
writer_err = pd.ExcelWriter(path=file_err, mode='w', engine='xlsxwriter')
writer_err_header = pd.ExcelWriter(path=file_header_err, mode='w', engine='xlsxwriter')

# 从文件列表遍历读取文件
for file_idx in range(len(file_name)):
    df_fcst = pd.read_excel(file_name[file_idx], dtype=str)
    # 判断header是否符合标准
    header_result = df_fcst.columns[0:34] == header_std
    if not header_result.all():
        row_data_err = [file_name[file_idx], 'Header Exception']
        data_err_header.append(row_data_err)
    else:
        # 符合header标准，遍历文件的所有行，验证行数据
        for idx_row in range(len(df_fcst[header_std[0]])):
            row_data = []  # 临时存放符合标准的idx_row行的数据
            row_data_err = []  # 临时存放错误的idx_row行的数据
            # 判断idx_row的前两列是否为空
            if not is_null(idx_row):
                add_error_log_and_data(file_name[file_idx], row_data_err, data_err, 'Issue/Dept Null Exception', idx_row)
            # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
            elif not calc_total_budget(idx_row):
                add_error_log_and_data(file_name[file_idx], row_data_err, data_err, 'TotalBudget Exception', idx_row)
            # 判断idx_row的Issue列是否含有多余值（Issue列值除【文件名内包含的issue值和Act】，其余都视为多余值）
            elif not is_duplicate(file_name[file_idx], idx_row):
                add_error_log_and_data(file_name[file_idx], row_data_err, data_err, 'Issue Exception', idx_row)
            # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
            elif not validate(idx_row):
                add_error_log_and_data(file_name[file_idx], row_data_err, data_err, 'Working/NonWorking Exception', idx_row)
            # 符合所有标准的数据格式化后写入Summary文件
            else:
                for idx_col in range(len(header_std)):
                    to_upper(idx_row, 5, 'Branding')
                    to_upper(idx_row, 5, 'NonBranding')
                    to_upper(idx_row, 6, 'Working')
                    to_upper(idx_row, 6, 'NonWorking')
                    to_float(idx_col, idx_row)
                    row_data.append(df_fcst[header_std[idx_col]][idx_row])
                data_new.append(row_data)
# 将异常header写入Error Header excel
frame_header_err = pd.DataFrame(data_err_header, columns=err_header_err)
frame_header_err.to_excel(writer_err_header, index=False, header=True)
to_format(writer_err_header, frame_header_err, [0, 1], [], [])
# 将异常数据写入Error excel
frame_err = pd.DataFrame(data_err, columns=header_err)
frame_err.to_excel(writer_err, index=False, header=True)
to_format(writer_err, frame_err, [0, 1, 2], [18, 19, 22, 35], [20, 21])
# 将标准数据写入Summary excel
frame_new = pd.DataFrame(data_new, columns=header_std)
frame_new.to_excel(writer_new, index=False, header=True)
to_format(writer_new, frame_new, [], [15, 16, 19, 32], [17, 18])
# 计算程序耗时
cost = time.time() - start
print(cost)