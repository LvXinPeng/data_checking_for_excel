import time
import pandas as pd
import os
import shutil
import datetime


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
    if str(df[header[16]][row_idx]).strip() not in ['Y', 'y', 'N', 'n']:
        is_ssmc = False
    return is_ssmc


def is_branding(df, header, row_idx):
    is_b = False
    if str(df[header[5]][row_idx]).replace(' ', '').upper() in ['BRANDING', 'NONBRANDING']:
        is_b = True
    return is_b


def calc_total_budget(df, header, row_idx):
    temp = []
    for col_idx in range(12):
        if str(df[header[col_idx + 20]][row_idx]) == '':
            df[header[col_idx + 20]][row_idx] = 0.0
            temp.append(0.0)
        elif str(df[header[col_idx + 20]][row_idx]).replace(' ', '') == '':
            df[header[col_idx + 20]][row_idx] = 0.0
            temp.append(0.0)
        elif str(df[header[col_idx + 20]][row_idx]).replace('-', '').replace(' ', '') == '':
            df[header[col_idx + 20]][row_idx] = 0.0
            temp.append(0.0)
        elif str(df[header[col_idx + 20]][row_idx]) == 'nan':
            df[header[col_idx + 20]][row_idx] = 0.0
            temp.append(0.0)
        else:
            try:
                temp.append(float(df[header[col_idx + 20]][row_idx]))
            except ValueError:
                temp.append(0.0)
    total_budget = sum(list(map(float, temp)))
    if not is_number(str(df[header[32]][row_idx])) or str(df[header[32]][row_idx]) in ['', ' ', 'nan']:
        df[header[32]][row_idx] = 0
    is_eq = total_budget == float(df[header[32]][row_idx]) or abs(total_budget - float(df[header[32]][row_idx])) <= 1
    return is_eq


def validate(df, header, row_idx):
    validation_sale = False
    validation_cate = False
    validation_acti = False
    for col_idx in range(10):
        to_replace(df, header, col_idx, row_idx)
    if str(df[header[6]][row_idx]).replace(' ', '').upper() == 'NONWORKING' \
            and str(df[header[7]][row_idx]).replace(' ', '').upper() \
            not in str(list(['Awareness', 'Consideration', 'Opinion'])).strip().upper():
        df[header[6]][row_idx] = 'NonWorking'
        validation_sale = True
        if str(df[header[8]][row_idx]).replace(' ', '').upper() in str(list(['Agency Fee'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() in str(list(
                    ['Media Agency', 'Event Agency', 'Creative production Agency', 'Social Agency', 'CRM Agency',
                     'PR Agency', 'Experimental Agency', 'Digital production Agency'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in str(list(['Marketing Audits'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in str(list(['Marketing Audits'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() in str(list(['Production'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() in str(list(
                    ['Production cost (origination)', 'Production cost (adaption, transcation, repurposing)',
                     'SEO'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in str(list(['Market Intelligence'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() in str(
                    list(['Market Research (consumer insight)', 'Business Intelligence'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in str(list(['Infrastructure Development'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in str(list(['MKT system development'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() in str(list(['Meetings'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() in str(
                    list(['Dealer Conference', 'Regional Meetings'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in str(list(['Operational Cost'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in str(list(['Car cost running (depreciation)', 'System (platform) maintenance', 'DTS'])).\
                    replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() in str(list(['Celebrity'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in str(list(['Celebrity'])).replace(' ', '').upper():
                validation_acti = True
    elif str(df[header[6]][row_idx]).replace(' ', '').upper() == 'WORKING' \
            and str(df[header[7]][row_idx]).replace(' ', '').upper() \
            in str(list(['Awareness', 'Consideration', 'Opinion'])).strip().upper():
        validation_sale = True
        if str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in str(list(['Traditional Media'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in str(list(['TV', 'Radio', 'Newspaper & Magazine', 'Cinema', 'OOH'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in str(list(['Digital Media'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in str(list(['Vertical', 'Vedio', 'News & Portal', 'SEM', 'Others'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in str(list(['Social Media'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in str(list(['Social Media'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() in str(list(['CRM'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in str(list(['CRM Campaign'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in str(list(['Call Center'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in str(list(['Consumer inbound & outbound'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() in str(list(['Event'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() in str(list(
                    ['Global Event', 'Test Drive', 'Autoshow', 'Product Display', 'Launch Campaign', 'Group Buy',
                     'Owner experience event', 'Seasonal Campaign', 'Cyber Campaign', 'Delivery Ceremony',
                     'Used car campaign', 'Plant Visit'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in str(list(['Public Relationship'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in str(list(['MY change communication', 'New product communication', 'Event communication'])).\
                    replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() in str(list(['POSM'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() in str(
                    list(['POSM', 'Customer Magazine', 'Dealer opening support'])).replace(' ', '').upper():
                validation_acti = True
        elif str(df[header[8]][row_idx]).replace(' ', '').upper() \
                in str(list(['Sponsorship'])).replace(' ', '').upper():
            validation_cate = True
            if str(df[header[9]][row_idx]).replace(' ', '').upper() \
                    in str(list(['Sponsorship'])).replace(' ', '').upper():
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
    return validation


def is_duplicate(df, f_path, f_name, row_idx):
    issue_name = f_name.split('.xlsx')[0].replace(f_path, '').split('&')[0]
    issue_names = [issue_name, 'Act', 'Actual']
    if df['Issue'][row_idx] not in issue_names:
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


def to_float(df, header, col_idx, row_idx):
    if 20 <= col_idx <= 32:
        if not is_number(str(df[header[col_idx]][row_idx])):
            df[header[col_idx]][row_idx] = float(0.0)
        try:
            df[header[col_idx]][row_idx] = float(df[header[col_idx]][row_idx])
        except ValueError:
            df[header[col_idx]][row_idx] = float(0.0)


def to_int(df, header, col_idx, row_idx):
    if 11 <= col_idx <= 15:
        if not is_number(str(df[header[col_idx]][row_idx])):
            df[header[col_idx]][row_idx] = 0
        try:
            df[header[col_idx]][row_idx] = int(df[header[col_idx]][row_idx])
        except ValueError:
            df[header[col_idx]][row_idx] = 0


def to_replace(df, header, col_idx, row_idx):
    if 0 <= col_idx <= 10:
        if str(df[header[col_idx]][row_idx]).find('_') != -1:
            df[header[col_idx]][row_idx] = str(df[header[col_idx]][row_idx]).replace('_', ' ')


def to_replace_null(df, header, col_idx, row_idx):
    if col_idx in [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15, 16, 18, 19, 33]:
        if str(df[header[col_idx]][row_idx]) in ["0", "nan"]:
            df[header[col_idx]][row_idx] = "N.A."
    if col_idx in [11, 12, 13, 14, 17, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32]:
        if str(df[header[col_idx]][row_idx]) in ["0", "nan"]:
            df[header[col_idx]][row_idx] = int(0)


def to_split_fcstact(df, header, f_path, row_idx, f_name):
    flag = True  # True表示Act，False表示FCST
    issue_name = f_name.split('.')[0].replace(f_path, '').split('&')[0]
    if str(df[header[0]][row_idx]).strip().upper() == issue_name.upper() \
            and str(df[header[17]][row_idx]) in ['nan', '0', '', ' ']:
        df[header[17]][row_idx] = 0
        flag = False
    # if str(df[header[0]][row_idx]) == 'Rebate':
    #     flag = True
    return flag


def to_dept(df, header, dept_name, row_idx):
    flag = False
    if str(df[header[1]][row_idx]).strip() == dept_name:
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
        if col in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget']:
            column_len = data_frame[col].astype(str).str.len().max()
        else:
            column_len = data_frame[col].astype(str).str.len().mean()
        # Setting the length if the column header is larger than the max column value length
        column_len = max(column_len, len(col)) + 2
        # set the column length
        worksheet.set_column(i, i, column_len)
    try:
        if len(data_frame.columns) < 35:
            for row_idx in range(len(data)):
                for col_idx in range(34):
                    if 20 <= col_idx <= 32:
                        worksheet.write(row_idx + 1, col_idx, data[row_idx][col_idx], budget_format)
        elif len(data_frame.columns) >= 35:
            for row_idx in range(len(data)):
                for col_idx in range(37):
                    if 23 <= col_idx <= 35:
                        worksheet.write(row_idx + 1, col_idx, data[row_idx][col_idx], budget_format)
    except IndexError:
        pass
    writer.save()
    writer.close()


def add_error_log_and_data(f_name, df, header, row_err_data, err_data, exception, row_idx):
    # 添加error log（三列）
    row_err_data.append(f_name)
    row_err_data.append(exception)
    row_err_data.append(row_idx + 2)
    # 添加error data（header列）
    for col_idx in range(len(header)):
        to_replace_null(df, header, col_idx, row_idx)
        to_float(df, header, col_idx, row_idx)
        to_int(df, header, col_idx, row_idx)
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
file_name_all = os.listdir('D:/ASP - Erin/Raw Data/')
file_name = []
for file_idx in range(len(file_name_all)):
    if file_name_all[file_idx].endswith('.xlsx'):
        file_name.append(path + file_name_all[file_idx])
data_new = []  # 存储ACT所有的标准数据
data_new_fcst = []  # 存储FCST所有的标准数据
data_err = []  # 存储ACT所有的异常数据
data_err_fcst = []  # 存储FCST所有的异常数据
data_err_header = []  # 存储所有的header异常数据
data_log = [] # 存储log里面的指标
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
today = str(datetime.date.today()).replace('-', '')
file_log = "D:/ASP - Erin/Log.xlsx"
file_new = "D:/ASP - Erin/Actual/Data&Log/Act_Summary_" + today + ".xlsx"
file_new_fcst = "D:/ASP - Erin/FCST/Data&Log/FCST_Summary_" + today + ".xlsx"
file_err = "D:/ASP - Erin/Actual/Data&Log/Act_Error_" + today + ".xlsx"
file_err_fcst = "D:/ASP - Erin/FCST/Data&Log/FCST_Error_" + today + ".xlsx"
file_header_err = "D:/ASP - Erin/Header Error/Header Error_" + today + ".xlsx"
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
    header_result = str(list(df_fcst.columns[0:34])).strip().upper() == str(header_std).upper()
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
            # df_fcst[header_std[0]][idx_row] = 'Actual'
            row_data = []  # 临时存放符合标准的idx_row行的数据
            row_data_err = []  # 临时存放错误的idx_row行的数据
            row_data_err_null = []  # 临时存放错误的idx_row行的数据
            row_data_err_calc = []  # 临时存放错误的idx_row行的数据
            row_data_err_dupl = []  # 临时存放错误的idx_row行的数据
            row_data_err_vali = []  # 临时存放错误的idx_row行的数据
            row_data_err_ssmc = []  # 临时存放错误的idx_row行的数据
            row_data_err_kpi = []  # 临时存放错误的idx_row行的数据
            row_data_err_brand = []  # 临时存放错误的idx_row行的数据
            # 判断idx_row的前两列是否为空
            if not is_null(df_fcst, header_std, idx_row):
                add_error_log_and_data(file_name[file_idx], df_fcst, header_std, row_data_err_null, data_err,
                                       'Issue/Dept Null Exception', idx_row)
            # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
            if not calc_total_budget(df_fcst, header_std, idx_row):
                add_error_log_and_data(file_name[file_idx], df_fcst, header_std, row_data_err_calc, data_err,
                                       'Budget Exception', idx_row)
            if not is_branding(df_fcst, header_std, idx_row):
                add_error_log_and_data(file_name[file_idx], df_fcst, header_std, row_data_err_brand, data_err,
                                       'Branding/NonBranding Exception', idx_row)
            # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
            if not validate(df_fcst, header_std, idx_row)[0]:
                add_error_log_and_data(file_name[file_idx], df_fcst, header_std, row_data_err_vali, data_err,
                                       'Sale Funnel Exception', idx_row)
            elif not validate(df_fcst, header_std, idx_row)[1]:
                add_error_log_and_data(file_name[file_idx], df_fcst, header_std, row_data_err_vali, data_err,
                                       'Category Exception', idx_row)
            elif not validate(df_fcst, header_std, idx_row)[2]:
                add_error_log_and_data(file_name[file_idx], df_fcst, header_std, row_data_err_vali, data_err,
                                       'Activity Type Exception', idx_row)
            if not is_ssm(df_fcst, header_std, idx_row):
                add_error_log_and_data(file_name[file_idx], df_fcst, header_std, row_data_err_ssmc, data_err,
                                       'SMM Exception', idx_row)
            if not validate_kpi(df_fcst, header_std, idx_row):
                add_error_log_and_data(file_name[file_idx], df_fcst, header_std, row_data_err_kpi, data_err,
                                       'KPI Exception', idx_row)
            # 符合所有标准的数据格式化后写入Summary文件
            if is_null(df_fcst, header_std, idx_row) and calc_total_budget(df_fcst, header_std, idx_row) \
                    and is_branding(df_fcst, header_std, idx_row) and validate(df_fcst, header_std, idx_row)[0] \
                    and validate(df_fcst, header_std, idx_row)[1] and validate(df_fcst, header_std, idx_row)[2] \
                    and is_ssm(df_fcst, header_std, idx_row) and validate_kpi(df_fcst, header_std, idx_row):
                # df_fcst[header_std[0]][idx_row] = 'Actual'
                for idx_col in range(len(header_std)):
                    to_replace(df_fcst, header_std, idx_col, idx_row)
                    to_simplify(df_fcst, header_std, idx_row)
                    to_replace_null(df_fcst, header_std, idx_col, idx_row)
                    to_upper(df_fcst, header_std, idx_row, 5, 'Branding')
                    to_upper(df_fcst, header_std, idx_row, 5, 'NonBranding')
                    to_upper(df_fcst, header_std, idx_row, 6, 'Working')
                    to_upper(df_fcst, header_std, idx_row, 6, 'NonWorking')
                    to_float(df_fcst, header_std, idx_col, idx_row)
                    to_int(df_fcst, header_std, idx_col, idx_row)
                    row_data.append(df_fcst[header_std[idx_col]][idx_row])
                    data_log.append(df_fcst[header_std[0]][idx_row])
                data_new.append(row_data)
# 将header异常数据写入Error excel
write_to_excel(data_err_header, err_header_err, writer_err_header, [0, 1], [], [])
# 将Actual异常数据写入Error excel
write_to_excel(data_err, header_err, writer_err, [0, 1, 2], [18, 19, 22, 35], [20, 21])
# 将FCST异常数据写入excel
write_to_excel(data_err_fcst, header_err, writer_err_fcst, [0, 1, 2], [18, 19, 22, 35], [20, 21])
# 将Actual标准数据写入excel
write_to_excel(data_new, header_std, writer_new, [], [15, 16, 19, 32], [17, 18])
# 将FCST标准数据写入Error excel
write_to_excel(data_new_fcst, header_std, writer_new_fcst, [], [15, 16, 19, 32], [17, 18])

path1 = "D:/ASP - Erin/FCST/Dept-Team/"
mkdir(path1)
# 读入待处理文件
file_summary = "D:/ASP - Erin/FCST/Data&Log/FCST_Summary_" + today + ".xlsx"
df_summary_fcst = pd.read_excel(file_summary, dtype=str)

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
for idx_row in range(len(df_summary_fcst[header_std[0]])):
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
    if str(df_summary_fcst[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKT':
        for idx_col in range(len(header_std)):
            to_float(df_summary_fcst, header_std, idx_col, idx_row)
            row_data_MKT.append(df_summary_fcst[header_std[idx_col]][idx_row])
        data_MKT.append(row_data_MKT)
    elif str(df_summary_fcst[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKTLAUNCH':
        for idx_col in range(len(header_std)):
            to_float(df_summary_fcst, header_std, idx_col, idx_row)
            row_data_MKT_Launch.append(df_summary_fcst[header_std[idx_col]][idx_row])
        data_MKT_Launch.append(row_data_MKT_Launch)
    elif str(df_summary_fcst[header_std[1]][idx_row]).replace(' ', '').upper() == 'APACCC':
        for idx_col in range(len(header_std)):
            to_float(df_summary_fcst, header_std, idx_col, idx_row)
            row_data_APAC_CC.append(df_summary_fcst[header_std[idx_col]][idx_row])
        data_APAC_CC.append(row_data_APAC_CC)
    elif str(df_summary_fcst[header_std[1]][idx_row]).replace(' ', '').upper() == 'CUSTOMERSERVICE':
        for idx_col in range(len(header_std)):
            to_float(df_summary_fcst, header_std, idx_col, idx_row)
            row_data_Customer_Service.append(df_summary_fcst[header_std[idx_col]][idx_row])
        data_Customer_Service.append(row_data_Customer_Service)
    elif str(df_summary_fcst[header_std[1]][idx_row]).replace(' ', '').upper() == 'DTS':
        for idx_col in range(len(header_std)):
            to_float(df_summary_fcst, header_std, idx_col, idx_row)
            row_data_DTS.append(df_summary_fcst[header_std[idx_col]][idx_row])
        data_DTS.append(row_data_DTS)
    elif str(df_summary_fcst[header_std[1]][idx_row]).replace(' ', '').upper() == 'MI':
        for idx_col in range(len(header_std)):
            to_float(df_summary_fcst, header_std, idx_col, idx_row)
            row_data_MI.append(df_summary_fcst[header_std[idx_col]][idx_row])
        data_MI.append(row_data_MI)
    elif str(df_summary_fcst[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
            df_summary_fcst[header_std[2]][idx_row]).replace(' ', '').upper() in str(list(['DMKT Central Team', 'Exhibition', 'East Region Team', 'North Region Team', 'South Region Team', 'West Region Team', 'ZheJiang Region Team'])).replace(' ', '').upper():
        for idx_col in range(len(header_std)):
            to_float(df_summary_fcst, header_std, idx_col, idx_row)
            row_data_Sales_MKT1.append(df_summary_fcst[header_std[idx_col]][idx_row])
        data_Sales_MKT1.append(row_data_Sales_MKT1)
    elif str(df_summary_fcst[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
            df_summary_fcst[header_std[2]][idx_row]).replace(' ', '').upper() in ['HK']:
        for idx_col in range(len(header_std)):
            to_float(df_summary_fcst, header_std, idx_col, idx_row)
            row_data_Sales_MKT2.append(df_summary_fcst[header_std[idx_col]][idx_row])
        data_Sales_MKT2.append(row_data_Sales_MKT2)
    elif str(df_summary_fcst[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
            df_summary_fcst[header_std[2]][idx_row]).replace(' ', '').upper() in ['INTERNALFLEETCAR']:
        for idx_col in range(len(header_std)):
            to_float(df_summary_fcst, header_std, idx_col, idx_row)
            row_data_Sales_MKT3.append(df_summary_fcst[header_std[idx_col]][idx_row])
        data_Sales_MKT3.append(row_data_Sales_MKT3)
    elif str(df_summary_fcst[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
            df_summary_fcst[header_std[2]][idx_row]).replace(' ', '').upper() in ['NATIONALSALES']:
        for idx_col in range(len(header_std)):
            to_float(df_summary_fcst, header_std, idx_col, idx_row)
            row_data_Sales_MKT4.append(df_summary_fcst[header_std[idx_col]][idx_row])
        data_Sales_MKT4.append(row_data_Sales_MKT4)
    elif str(df_summary_fcst[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
            df_summary_fcst[header_std[2]][idx_row]).replace(' ', '').upper() in ['NBDTEAM']:
        for idx_col in range(len(header_std)):
            to_float(df_summary_fcst, header_std, idx_col, idx_row)
            row_data_Sales_MKT5.append(df_summary_fcst[header_std[idx_col]][idx_row])
        data_Sales_MKT5.append(row_data_Sales_MKT5)
# 将标准数据写入Summary excel
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

path2 = "D:/ASP - Erin/Actual/Dept-Team/"
mkdir(path2)
# 读入待处理文件
file_summary2 = "D:/ASP - Erin/Actual/Data&Log/Act_Summary_" + today + ".xlsx"
df_summary_actual = pd.read_excel(file_summary2, dtype=str)

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
for idx_row in range(len(df_summary_actual[header_std[0]])):
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
    if str(df_summary_actual[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKT':
        for idx_col in range(len(header_std)):
            to_float(df_summary_actual, header_std, idx_col, idx_row)
            row_data_MKT.append(df_summary_actual[header_std[idx_col]][idx_row])
        data_MKT2.append(row_data_MKT)
    elif str(df_summary_actual[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKTLAUNCH':
        for idx_col in range(len(header_std)):
            to_float(df_summary_actual, header_std, idx_col, idx_row)
            row_data_MKT_Launch.append(df_summary_actual[header_std[idx_col]][idx_row])
        data_MKT_Launch2.append(row_data_MKT_Launch)
    elif str(df_summary_actual[header_std[1]][idx_row]).replace(' ', '').upper() == 'APACCC':
        for idx_col in range(len(header_std)):
            to_float(df_summary_actual, header_std, idx_col, idx_row)
            row_data_APAC_CC.append(df_summary_actual[header_std[idx_col]][idx_row])
        data_APAC_CC2.append(row_data_APAC_CC)
    elif str(df_summary_actual[header_std[1]][idx_row]).replace(' ', '').upper() == 'CUSTOMERSERVICE':
        for idx_col in range(len(header_std)):
            to_float(df_summary_actual, header_std, idx_col, idx_row)
            row_data_Customer_Service.append(df_summary_actual[header_std[idx_col]][idx_row])
        data_Customer_Service2.append(row_data_Customer_Service)
    elif str(df_summary_actual[header_std[1]][idx_row]).replace(' ', '').upper() == 'DTS':
        for idx_col in range(len(header_std)):
            to_float(df_summary_actual, header_std, idx_col, idx_row)
            row_data_DTS.append(df_summary_actual[header_std[idx_col]][idx_row])
        data_DTS2.append(row_data_DTS)
    elif str(df_summary_actual[header_std[1]][idx_row]).replace(' ', '').upper() == 'MI':
        for idx_col in range(len(header_std)):
            to_float(df_summary_actual, header_std, idx_col, idx_row)
            row_data_MI.append(df_summary_actual[header_std[idx_col]][idx_row])
        data_MI2.append(row_data_MI)
    elif str(df_summary_actual[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
            df_summary_actual[header_std[2]][idx_row]).replace(' ', '').upper() in str(list(['DMKT Central Team', 'Exhibition', 'East Region Team', 'North Region Team', 'South Region Team', 'West Region Team', 'ZheJiang Region Team'])).replace(' ', '').upper():
        for idx_col in range(len(header_std)):
            to_float(df_summary_actual, header_std, idx_col, idx_row)
            row_data_Sales_MKT1.append(df_summary_actual[header_std[idx_col]][idx_row])
        data_Sales_MKT12.append(row_data_Sales_MKT1)
    elif str(df_summary_actual[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
            df_summary_actual[header_std[2]][idx_row]).replace(' ', '').upper() in ['HK']:
        for idx_col in range(len(header_std)):
            to_float(df_summary_actual, header_std, idx_col, idx_row)
            row_data_Sales_MKT2.append(df_summary_actual[header_std[idx_col]][idx_row])
        data_Sales_MKT22.append(row_data_Sales_MKT2)
    elif str(df_summary_actual[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
            df_summary_actual[header_std[2]][idx_row]).replace(' ', '').upper() in ['INTERNALFLEETCAR']:
        for idx_col in range(len(header_std)):
            to_float(df_summary_actual, header_std, idx_col, idx_row)
            row_data_Sales_MKT3.append(df_summary_actual[header_std[idx_col]][idx_row])
        data_Sales_MKT32.append(row_data_Sales_MKT3)
    elif str(df_summary_actual[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
            df_summary_actual[header_std[2]][idx_row]).replace(' ', '').upper() in ['NATIONALSALES']:
        for idx_col in range(len(header_std)):
            to_float(df_summary_actual, header_std, idx_col, idx_row)
            row_data_Sales_MKT4.append(df_summary_actual[header_std[idx_col]][idx_row])
        data_Sales_MKT42.append(row_data_Sales_MKT4)
    elif str(df_summary_actual[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
            df_summary_actual[header_std[2]][idx_row]).replace(' ', '').upper() in ['NBDTEAM']:
        for idx_col in range(len(header_std)):
            to_float(df_summary_actual, header_std, idx_col, idx_row)
            row_data_Sales_MKT5.append(df_summary_actual[header_std[idx_col]][idx_row])
        data_Sales_MKT52.append(row_data_Sales_MKT5)
# 将标准数据写入Summary excel
write_to_excel(data_MKT2, header_std, writer_MKT2, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_MKT_Launch2, header_std, writer_MKT_Launch2, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_APAC_CC2, header_std, writer_APAC_CC2, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Customer_Service2, header_std, writer_Customer_Service2, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_DTS2, header_std, writer_DTS2, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_MI2, header_std, writer_MI2, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT12, header_std, writer_Sales_MKT12, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT22, header_std, writer_Sales_MKT22, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT32, header_std, writer_Sales_MKT32, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT42, header_std, writer_Sales_MKT42, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT52, header_std, writer_Sales_MKT52, [], [15, 16, 19, 32], [17, 18])

# 计算程序耗时
cost = time.time() - start
print(cost)
# 生成log
try:
    log_time = []
    log_col = ['LastUpdatedTime', 'Issue']
    data_log_unique = list(set(data_log))
    if 'Actual' in data_log_unique:
        data_log_unique.remove('Actual')
    if 'Act' in data_log_unique:
        data_log_unique.remove('Act')
    for idx in range(len(data_log_unique)):
        log_time.append([str(datetime.datetime.now()), data_log_unique[idx]])
    writer_log = pd.ExcelWriter(path=file_log, mode='w', engine='xlsxwriter')
    frame_log = pd.DataFrame(log_time, columns=log_col)
    frame_log.to_excel(writer_log, index=False, header=True)
    writer_log.save()
    writer_log.close()
except KeyError or IndexError:
    pass
print("Process finished with exit code 0")
