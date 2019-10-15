import stat
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


def is_life_cycle(df, header, row_idx):
    is_life = False
    if str(df[header[4]][row_idx]).replace(' ', '').upper() \
            in str(list(['Launch', 'Sustain', 'MY Change'])).replace(' ', '').upper():
        is_life = True
    return is_life


def is_carline(df, header, row_idx):
    is_car = False
    if str(df[header[3]][row_idx]).replace(' ', '').upper() \
            in str(list(['XC90 (V526)', 'XC60 (K426)', 'XC40 (V316)', 'XC40 (K316)', 'S90L', 'S60', 'S60L', 'V90CC',
                         'V60', 'V60CC', 'V40', 'V40CC', 'S90 Excellence', 'XC90 Excellence', 'Polestar',
                         'S90 Ambiance', 'XC90 Ambiance', 'Mix Carline', 'Customer Service'])).replace(' ', '').upper():
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


def calc_ytd_budget(df, header, row_idx, month):
    temp = []
    for col_idx in range(month):
        try:
            temp.append(float(df[header[col_idx + 20]][row_idx]))
        except ValueError:
            temp.append(0.0)
    total_ytd_budget = sum(list(map(float, temp)))
    return total_ytd_budget


def get_month(f_path):
    df = pd.read_excel(f_path, dtype=str)
    month = int(list(str(df['Issue'][0]).replace(' ', '').split('+')[0])[-1])
    return month


def validate_team(df, header, row_idx):
    validation = False
    if str(df[header_std[1]][row_idx]).replace(' ', '').upper() == 'MKT' \
            and str(df[header[2]][row_idx]).replace(' ', '').upper() \
            in ['TRADITIONALMEDIATEAM', 'DIGITALMEDIATEAM', 'SOCIALMEDIATEAM', 'CRMTEAM', 'EVENTTEAM',
                'PUBLICRELATIONSHIP',
                'CREATIVEPRODUCTION', 'DIGITALPRODUCTION', 'STRATEGYTEAM', 'LAUNCHTEAM', 'CSMKT']:
        validation = True
    elif str(df[header[1]][row_idx]).replace(' ', '').upper() == 'MKTLAUNCH' \
            and str(df[header[2]][row_idx]).replace(' ', '').upper() == 'LAUNCHTEAM':
        validation = True
    elif str(df[header[1]][row_idx]).replace(' ', '').upper() == 'APACCCBRANDING' \
            and str(df[header[2]][row_idx]).replace(' ', '').upper() == 'APACCCBRANDING':
        validation = True
    elif str(df[header[1]][row_idx]).replace(' ', '').upper() == 'APACCC' \
            and str(df[header[2]][row_idx]).replace(' ', '').upper() == 'APACCC':
        validation = True
    elif str(df[header[1]][row_idx]).replace(' ', '').upper() == 'CUSTOMERSERVICE' \
            and str(df[header[2]][row_idx]).replace(' ', '').upper() \
            in ['CSEVENT', 'CUSTOMERCARE', 'SERVICEOFFERS', 'SERVICEEFFICIENCY', 'REGIONALCOLLABORATION']:
        validation = True
    elif str(df[header[1]][row_idx]).replace(' ', '').upper() == 'DTS' \
            and str(df[header[2]][row_idx]).replace(' ', '').upper() == 'DTS':
        validation = True
    elif str(df[header[1]][row_idx]).replace(' ', '').upper() == 'MI' \
            and str(df[header[2]][row_idx]).replace(' ', '').upper() == 'MI':
        validation = True
    elif str(df[header[1]][row_idx]).replace(' ', '').upper() in ['SALESMKT'] \
            and str(df[header[2]][row_idx]).replace(' ', '').upper() \
            in ['DMKTCENTRALTEAM', 'EXHIBITION', 'EASTREGIONTEAM', 'NORTHREGIONTEAM', 'SOUTHREGIONTEAM',
                'WESTREGIONTEAM',
                'ZHEJIANGREGIONTEAM', 'HK', 'FLEETTEAM', 'USEDCARTEAM', 'NATIONALSALES', 'INTERNALFLEETCAR',
                'FINANCIALSERVICETEAM']:
        validation = True
    return validation


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
                    in str(list(['Car cost running (depreciation)', 'System (platform) maintenance', 'DTS'])). \
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
                    in str(list(['MY change communication', 'New product communication', 'Event communication'])). \
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
    if str(df[header[16]][row_idx]).strip().upper() == 'N' \
            and (str(df[header[11]][row_idx]) not in ['0', 'nan', '/', '-']
                 or str(df[header[12]][row_idx]) not in ['0', 'nan', '/', '-']
                 or str(df[header[13]][row_idx]) not in ['0', 'nan', '/', '-']
                 or str(df[header[14]][row_idx]) not in ['0', 'nan', '/', '-']):
        validation = False
    return validation


def is_duplicate(df, issue_name, row_idx):
    issue_name_upper = str(issue_name).replace(' ', '').upper()
    issue_names = [issue_name_upper, 'ACT', 'ACTUAL']
    if str(df['Issue'][row_idx]).replace(' ', '').upper() in issue_names:
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
    if 20 <= col_idx <= 32 or 34 <= col_idx <= 38:
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
    if 0 <= col_idx <= 10:
        if str(df[header[col_idx]][row_idx]).find('_') != -1:
            df[header[col_idx]][row_idx] = str(df[header[col_idx]][row_idx]).replace('_', ' ')


def to_replace_null(df, header, col_idx, row_idx):
    if col_idx in [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15, 16, 18, 19, 33]:
        if str(df[header[col_idx]][row_idx]).replace(' ', '') in ["0", "nan", '-', '/', '']:
            df[header[col_idx]][row_idx] = "N.A."
    if col_idx in [11, 12, 13, 14, 17, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32]:
        if str(df[header[col_idx]][row_idx]).replace(' ', '') in ["0", "nan", '-', '/', '']:
            df[header[col_idx]][row_idx] = int(0)


def is_actual(df, header, issue_name, row_idx):
    flag = False
    if str(df[header[0]][row_idx]).replace(' ', '').upper() == str(issue_name).replace(' ', '').upper() \
            and str(df[header[17]][row_idx]) != '0':
        flag = True
    elif str(df[header[0]][row_idx]).replace(' ', '').upper() in ['ACT', 'ACTUAL']:
        flag = True
    return flag


def is_actual_in_fcst(df, header, issue_name, row_idx):
    flag = False
    if str(df[header[0]][row_idx]).replace(' ', '').upper() == str(issue_name).replace(' ', '').upper() \
            and str(df[header[17]][row_idx]) != '0':
        flag = True
    return flag


def is_actual_only(df, header, row_idx):
    flag = False
    if str(df[header[0]][row_idx]).replace(' ', '').upper() in ['ACT', 'ACTUAL']:
        flag = True
    return flag


def is_forecast(df, header, issue_name, row_idx):
    flag = False
    if str(df[header[0]][row_idx]).replace(' ', '').upper() == str(issue_name).replace(' ', '').upper():
        flag = True
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
        if col in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
                   'Total Budget', 'Q1', 'Q2', 'Q3', 'Q4', 'YTD']:
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
        elif len(data_frame.columns) == 39:
            for row_idx in range(len(data)):
                for col_idx in range(39):
                    if 20 <= col_idx <= 32 or 34 <= col_idx <= 38:
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
        # to_replace_null(df, header, col_idx, row_idx)
        # to_float(df, header, col_idx, row_idx)
        # to_int(df, header, col_idx, row_idx)
        row_err_data.append(df[header[col_idx]][row_idx])
    err_data.append(row_err_data)


def write_to_excel(data, header, writer, err_col, red_col, gray_col):
    frame = pd.DataFrame(data, columns=header)
    frame.to_excel(writer, index=False, header=True)
    to_format(writer, frame, data, err_col, red_col, gray_col)


def is_writeable(path, check_parent=False):
    if os.access(path, os.F_OK) and os.access(path, os.W_OK):
        return True
    if os.access(path, os.F_OK) and not os.access(path, os.W_OK):
        return False
    if check_parent is False:
        return False
    parent_dir = os.path.dirname(path)
    if not os.access(parent_dir, os.F_OK):
        return False
    return os.access(parent_dir, os.W_OK)


def add_files(path, files, flag):
    folder = os.listdir(path)
    for file_idx in range(len(folder)):
        if folder[file_idx].endswith('.xlsx'):
            if flag == 'readable' and not is_writeable(path + folder[file_idx], check_parent=False):
                files.append(path + folder[file_idx])
            elif flag == 'writeable' and is_writeable(path + folder[file_idx], check_parent=False):
                files.append(path + folder[file_idx])


def add_summary(file_name, header, data):
    df_summary = pd.read_excel(file_name, dtype=str)
    for idx_row in range(len(df_summary[header[0]])):
        row_data = []
        for idx_col in range(len(header)):
            to_float(df_summary, header, idx_col, idx_row)
            to_int(df_summary, header, idx_col, idx_row)
            row_data.append(df_summary[header[idx_col]][idx_row])
        data.append(row_data)
    os.chmod(file_name, stat.S_IWRITE)


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
                   'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)', 'Q1', 'Q2', 'Q3', 'Q4', 'YTD']
# 异常数据excel header
header_err = ['File Name', 'Exception Type', 'Index', 'Issue', 'Dept', 'Team', 'Carline', 'Lifecycle',
              'Branding/NonBranding', 'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type', 'Activity',
              'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order', 'KPI Others', 'SMM Campaign Code (Y/N)',
              'SC NO.', ' SC Name', 'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct',
              'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
# 异常header excel header
err_header_err = ['File Name', 'Exception Type']
# 读入待处理文件
path_fcst = "D:/ASP - Erin/FCST/Data&Log/"
path_actual = "D:/ASP - Erin/Actual/Data&Log/"
path_header = "D:/ASP - Erin/Header Error/Files/"
file_readable_name_actual = []
file_readable_name_fcst = []
file_writeable_name = []
add_files(path_fcst, file_readable_name_fcst, 'readable')
add_files(path_fcst, file_writeable_name, 'writeable')
add_files(path_actual, file_readable_name_actual, 'readable')
add_files(path_actual, file_writeable_name, 'writeable')
add_files(path_header, file_writeable_name, 'writeable')

data_new_actual = []  # 存储ACT所有的标准数据
data_new_fcst = []  # 存储FCST所有的标准数据
data_err_actual = []  # 存储ACT所有的异常数据
data_err_fcst = []  # 存储FCST所有的异常数据
data_err_header = []  # 存储所有的header异常数据
# # 创建存放excel的文件夹
file_path_a = "D:/ASP - Erin/FCST/Data&Log/"
file_path_b = "D:/ASP - Erin/Header Error/"
file_path_c = "D:/ASP - Erin/Header Error/Files/"
file_path_d = "D:/ASP - Erin/Actual/Data&Log/"
# mkdir(file_path_a)
# mkdir(file_path_b)
# mkdir(file_path_c)
# mkdir(file_path_d)
# 几个生成的excel路径及文件名
today = str(datetime.date.today()).replace('-', '')
file_log = "D:/ASP - Erin/Log.xlsx"
file_new_actual = "D:/ASP - Erin/Actual/Data&Log/Act_Summary.xlsx"
file_new_fcst = "D:/ASP - Erin/FCST/Data&Log/FCST_Summary.xlsx"
file_err_actual = "D:/ASP - Erin/Actual/Data&Log/Act_Error.xlsx"
file_err_fcst = "D:/ASP - Erin/FCST/Data&Log/FCST_Error.xlsx"
file_header_err = "D:/ASP - Erin/Header Error/Header Error.xlsx"
# 创建几个excel的Writer引擎
writer_new_actual = pd.ExcelWriter(path=file_new_actual, mode='w', engine='xlsxwriter')
writer_new_fcst = pd.ExcelWriter(path=file_new_fcst, mode='w', engine='xlsxwriter')
writer_err_actual = pd.ExcelWriter(path=file_err_actual, mode='w', engine='xlsxwriter')
writer_err_fcst = pd.ExcelWriter(path=file_err_fcst, mode='w', engine='xlsxwriter')
writer_err_header = pd.ExcelWriter(path=file_header_err, mode='w', engine='xlsxwriter')

add_summary(file_readable_name_fcst[0], header_std_plus, data_new_fcst)
add_summary(file_readable_name_actual[0], header_std_plus, data_new_actual)

# 取Issue值
issue_val = pd.read_excel(file_log, dtype=str)['Issue'][0]
cur_month = get_month(file_log)
# 从文件列表遍历读取文件
for file_idx in range(len(file_writeable_name)):
    df_asp = pd.read_excel(file_writeable_name[file_idx], dtype=str)
    # 判断header是否符合标准
    header_result = str(list(df_asp.columns[0:34])).strip() == str(header_std)
    if not header_result:
        row_data_err = [file_writeable_name[file_idx], 'Header Exception']
        data_err_header.append(row_data_err)
        try:
            shutil.copy(file_writeable_name[file_idx], file_path_c)
        except shutil.SameFileError:
            pass
    else:
        for idx_row in range(len(df_asp[header_std[0]])):
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
            # 符合header标准，遍历文件的所有行，验证行数据 data checking
            if is_forecast(df_asp, header_std, issue_val, idx_row):
                row_data = []  # 临时存放符合标准的idx_row行的数据
                row_data_err = []  # 临时存放错误的idx_row行的数据
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
                # 判断idx_row的前两列是否为空
                if not is_null(df_asp, header_std, idx_row):
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_null_fcst,
                                           data_err_fcst,
                                           'Issue/Dept Null Exception', idx_row)
                if not is_carline(df_asp, header_std, idx_row):
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_carl_fcst,
                                           data_err_fcst,
                                           'Carline Exception', idx_row)
                if not is_life_cycle(df_asp, header_std, idx_row):
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_life_fcst,
                                           data_err_fcst,
                                           'Lifecycle Exception', idx_row)
                # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                if not calc_total_budget(df_asp, header_std, idx_row):
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_calc_fcst,
                                           data_err_fcst,
                                           'TotalBudget Exception', idx_row)
                if not is_branding(df_asp, header_std, idx_row):
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_brand_fcst,
                                           data_err_fcst,
                                           'Branding/NonBranding Exception', idx_row)
                if not is_num_kpi(df_asp, header_std, idx_row)[0]:
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_kpin_fcst,
                                           data_err_fcst,
                                           'KPI Prospects Exception', idx_row)
                elif not is_num_kpi(df_asp, header_std, idx_row)[1]:
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_kpin_fcst,
                                           data_err_fcst,
                                           'KPI Leads Exception', idx_row)
                elif not is_num_kpi(df_asp, header_std, idx_row)[2]:
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_kpin_fcst,
                                           data_err_fcst,
                                           'KPI Inquiry Exception', idx_row)
                elif not is_num_kpi(df_asp, header_std, idx_row)[3]:
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_kpin_fcst,
                                           data_err_fcst,
                                           'KPI Order Exception', idx_row)
                if not is_duplicate(df_asp, issue_val, idx_row):
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_dupl_fcst,
                                           data_err_fcst,
                                           'Issue Exception', idx_row)
                # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
                if not validate_team(df_asp, header_std, idx_row):
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_team_fcst,
                                           data_err_fcst,
                                           'Team Exception', idx_row)
                if not validate(df_asp, header_std, idx_row)[0]:
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_vali_fcst,
                                           data_err_fcst,
                                           'Sale Funnel Exception', idx_row)
                elif not validate(df_asp, header_std, idx_row)[1]:
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_vali_fcst,
                                           data_err_fcst,
                                           'Category Exception', idx_row)
                elif not validate(df_asp, header_std, idx_row)[2]:
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_vali_fcst,
                                           data_err_fcst,
                                           'Activity Type Exception', idx_row)
                if not is_ssm(df_asp, header_std, idx_row):
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_ssmc_fcst,
                                           data_err_fcst,
                                           'SMM Exception', idx_row)
                if not validate_kpi(df_asp, header_std, idx_row):
                    add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_kpi_fcst,
                                           data_err_fcst,
                                           'KPI Exception', idx_row)
                # 符合所有标准的数据格式化后写入Summary文件
                if is_null(df_asp, header_std, idx_row) and calc_total_budget(df_asp, header_std, idx_row) \
                        and is_branding(df_asp, header_std, idx_row) and validate(df_asp, header_std, idx_row)[0] \
                        and validate(df_asp, header_std, idx_row)[1] and validate(df_asp, header_std, idx_row)[2] \
                        and is_ssm(df_asp, header_std, idx_row) and validate_kpi(df_asp, header_std, idx_row) \
                        and is_duplicate(df_asp, issue_val, idx_row) \
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
                    row_data.append(calc_ytd_budget(df_asp, header_std, idx_row, cur_month))
                    data_new_fcst.append(row_data)
            if is_actual(df_asp, header_std, issue_val, idx_row):
                row_data = []  # 临时存放符合标准的idx_row行的数据
                row_data_err = []  # 临时存放错误的idx_row行的数据
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
                if is_actual_only(df_asp, header_std, idx_row):
                    # 判断idx_row的前两列是否为空
                    if not is_null(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_null,
                                               data_err_actual,
                                               'Issue/Dept Null Exception', idx_row)
                    if not is_carline(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_carl,
                                               data_err_actual,
                                               'Carline Exception', idx_row)
                    if not is_life_cycle(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_life,
                                               data_err_actual,
                                               'Lifecycle Exception', idx_row)
                    # 判断idx_row的12个月的budget计算总和是否和Total Budget列相等
                    if not calc_total_budget(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_calc,
                                               data_err_actual,
                                               'Budget Exception', idx_row)
                    if not is_branding(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_brand,
                                               data_err_actual,
                                               'Branding/NonBranding Exception', idx_row)
                    if not is_duplicate(df_asp, issue_val, idx_row):
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_dupl,
                                               data_err_actual,
                                               'Issue Exception', idx_row)
                    # 判断idx_row的Working/NonWorking列与某几列的关联关系是否符合标准
                    if not is_num_kpi(df_asp, header_std, idx_row)[0]:
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_kpin,
                                               data_err_actual,
                                               'KPI Prospects Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[1]:
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_kpin,
                                               data_err_actual,
                                               'KPI Leads Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[2]:
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_kpin,
                                               data_err_actual,
                                               'KPI Inquiry Exception', idx_row)
                    elif not is_num_kpi(df_asp, header_std, idx_row)[3]:
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_kpin,
                                               data_err_actual,
                                               'KPI Order Exception', idx_row)
                    if not validate_team(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_team,
                                               data_err_actual,
                                               'Team Exception', idx_row)
                    if not validate(df_asp, header_std, idx_row)[0]:
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_vali,
                                               data_err_actual,
                                               'Sale Funnel Exception', idx_row)
                    elif not validate(df_asp, header_std, idx_row)[1]:
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_vali,
                                               data_err_actual,
                                               'Category Exception', idx_row)
                    elif not validate(df_asp, header_std, idx_row)[2]:
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_vali,
                                               data_err_actual,
                                               'Activity Type Exception', idx_row)
                    if not is_ssm(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_ssmc,
                                               data_err_actual,
                                               'SMM Exception', idx_row)
                    if not validate_kpi(df_asp, header_std, idx_row):
                        add_error_log_and_data(file_writeable_name[file_idx], df_asp, header_std, row_data_err_kpi,
                                               data_err_actual,
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
                        and is_duplicate(df_asp, issue_val, idx_row):
                    for idx_col in range(len(header_std)):
                        row_data.append(df_asp[header_std[idx_col]][idx_row])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[0])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[1])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[2])
                    row_data.append(calc_q_budget(df_asp, header_std, idx_row)[3])
                    row_data.append(calc_ytd_budget(df_asp, header_std, idx_row, cur_month))
                    data_new_actual.append(row_data)
# 将header异常数据写入Error excel
write_to_excel(data_err_header, err_header_err, writer_err_header, [0, 1], [], [])
# 将Actual异常数据写入Error excel
write_to_excel(data_err_actual, header_err, writer_err_actual, [0, 1, 2], [18, 19, 22, 35], [20, 21])
# 将FCST异常数据写入excel
write_to_excel(data_err_fcst, header_err, writer_err_fcst, [0, 1, 2], [18, 19, 22, 35], [20, 21])
# 将Actual标准数据写入excel
write_to_excel(data_new_actual, header_std_plus, writer_new_actual, [], [15, 16, 19, 32], [17, 18])
# 将FCST标准数据写入Error excel
write_to_excel(data_new_fcst, header_std_plus, writer_new_fcst, [], [15, 16, 19, 32], [17, 18])
os.chmod(file_new_actual, stat.S_IREAD)
os.chmod(file_new_fcst, stat.S_IREAD)
cost = time.time() - start
print(cost)
print("Process finished with exit code 0")