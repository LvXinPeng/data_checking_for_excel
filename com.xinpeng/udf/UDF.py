import os
import pandas as pd


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
            except:
                temp.append(0.0)
    total_budget = sum(list(map(float, temp)))
    if not is_number(str(df[header[32]][row_idx])) or str(df[header[32]][row_idx]) in ['', ' ', 'nan']:
        df[header[32]][row_idx] = 0
    is_eq = total_budget == float(df[header[32]][row_idx]) or abs(total_budget - float(df[header[32]][row_idx])) <= 1
    return is_eq


def validate(df, header, row_idx):
    validation = False
    for col_idx in range(10):
        to_replace(df, header, col_idx, row_idx)
    if str(df[header[6]][row_idx]).strip().upper() == 'NONWORKING' and str(df[header[7]][row_idx]).strip() not in ['Awareness', 'Consideration', 'Opinion']:
        if str(df[header[8]][row_idx]).strip() in ['Agency Fee']:
            if str(df[header[9]][row_idx]).strip() in ['Media Agency', 'Event Agency', 'Creative production Agency', 'Social Agency', 'CRM Agency', 'PR Agency', 'Experimental Agency', 'Digital production Agency']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Marketing Audits']:
            if str(df[header[9]][row_idx]).strip() in ['Marketing Audits']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Production']:
            if str(df[header[9]][row_idx]).strip() in ['Production cost (origination)', 'Production cost (adaption, transcation, repurposing)', 'SEO']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Market Intelligence']:
            if str(df[header[9]][row_idx]).strip() in ['Market Research (consumer insight)', 'Business Intelligence']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Infrastructure Development']:
            if str(df[header[9]][row_idx]).strip() in ['MKT system development']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Meetings']:
            if str(df[header[9]][row_idx]).strip() in ['Dealer Conference', 'Regional Meetings']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Operational Cost']:
            if str(df[header[9]][row_idx]).strip() in ['Car cost running (depreciation)', 'System (platform) maintenance', 'DTS']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Celebrity']:
            if str(df[header[9]][row_idx]).strip() in ['Celebrity']:
                validation = True
    elif str(df[header[6]][row_idx]).strip().upper() == 'WORKING' and str(df[header[7]][row_idx]).strip() in ['Awareness', 'Consideration', 'Opinion']:
        if str(df[header[8]][row_idx]).strip() in ['Traditional Media']:
            if str(df[header[9]][row_idx]).strip() in ['TV', 'Radio', 'Newspaper & Magazine', 'Cinema', 'OOH']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Digital Media']:
            if str(df[header[9]][row_idx]).strip() in ['Vertical', 'Vedio', 'News & Portal', 'SEM', 'Others']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Social Media']:
            if str(df[header[9]][row_idx]).strip() in ['Social Media']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['CRM']:
            if str(df[header[9]][row_idx]).strip() in ['CRM Campaign']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Call Center']:
            if str(df[header[9]][row_idx]).strip() in ['Consumer inbound & outbound']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Event']:
            if str(df[header[9]][row_idx]).strip() in ['Global Event', 'Test Drive', 'Autoshow', 'Product Display', 'Launch Campaign', 'Group Buy', 'Owner experience event', 'Seasonal Campaign', 'Cyber Campaign', 'Delivery Ceremony', 'Used car campaign', 'Plant Visit']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Public Relationship']:
            if str(df[header[9]][row_idx]).strip() in ['MY change communication', 'New product communication', 'Event communication']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['POSM']:
            if str(df[header[9]][row_idx]).strip() in ['POSM', 'Customer Magazine', 'Dealer opening support']:
                validation = True
        elif str(df[header[8]][row_idx]).strip() in ['Sponsorship']:
            if str(df[header[9]][row_idx]).strip() in ['Sponsorship']:
                validation = True
    return validation


def validate_kpi(df, header, row_idx):
    validation = True
    if str(df[header[16]][row_idx]).strip().upper() == 'Y' and str(df[header[11]][row_idx]) in ['0', 'nan', '/', '-'] and str(
        df[header[14]][row_idx]) in ['0', 'nan', '/', '-'] and str(df[header[13]][row_idx]) in [
        '0', 'nan', '/', '-'] and str(df[header[12]][row_idx]) in ['0', 'nan', '/', '-']:
        validation = False
    # if str(df[header[16]][row_idx]).strip().upper() == 'Y' and is_number(str(df[header[11]][row_idx])):
    #     validation = True
    return validation


def is_duplicate(df, path, f_name, row_idx):
    issue_name = f_name.split('.xlsx')[0].replace(path, '').split('&')[0]
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
        except:
            pass


def to_int(df, header, col_idx, row_idx):
    if 11 <= col_idx <= 15:
        if not is_number(str(df[header[col_idx]][row_idx])):
            df[header[col_idx]][row_idx] = 0
        try:
            df[header[col_idx]][row_idx] = int(df[header[col_idx]][row_idx])
        except:
            pass


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


def to_split_fcstact(df, header, path, row_idx, f_name):
    flag = True  # True表示Act，False表示FCST
    issue_name = f_name.split('.')[0].replace(path, '').split('&')[0]
    if str(df[header[0]][row_idx]).strip().upper() == issue_name.upper() and str(df[header[17]][row_idx]) in ['nan', '0', '',  ' ']:
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


def add_error_log_and_data(f_name, df, header,  row_err_data, err_data, exception, row_idx):
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