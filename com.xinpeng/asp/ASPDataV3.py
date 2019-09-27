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
    is_eq = False
    if df_fcst[header_std[20]][row_idx] == '' or df_fcst[header_std[21]][row_idx] == '' or df_fcst[header_std[22]][row_idx] == '' or df_fcst[header_std[23]][row_idx] == '' or df_fcst[header_std[24]][row_idx] == '' or df_fcst[header_std[25]][row_idx] == '' or df_fcst[header_std[26]][row_idx] == '' or df_fcst[header_std[27]][row_idx] == '' or df_fcst[header_std[28]][row_idx] == '' or df_fcst[header_std[29]][row_idx] == '' or df_fcst[header_std[30]][row_idx] == '' or df_fcst[header_std[31]][row_idx] == '' or df_fcst[header_std[32]][row_idx] == '':
        return is_eq
    else:
        for col_idx in range(12):
            if str(df_fcst[header_std[col_idx + 20]][row_idx]).replace(' ', '') == '' or str(df_fcst[header_std[col_idx + 20]][row_idx]).replace('-', '') == '':
                temp.append(0)
            else:
                temp.append(df_fcst[header_std[col_idx + 20]][row_idx])
        total_budget = sum(list(map(float, temp)))
        is_eq = total_budget == float(df_fcst[header_std[32]][row_idx])
    return is_eq


def validate(row_idx):
    validation = True
    if str(df_fcst[header_std[6]][row_idx]).upper() == 'NONWORKING' and df_fcst[header_std[7]][row_idx] != "0":
        if df_fcst[header_std[8]][row_idx] not in ['Agency Fee', 'Market Intelligence','Marketing Audits','Production','Infrastructure Development','Meetings','Operational Cost','Celebrity']:
            if df_fcst[header_std[9]][row_idx] not in ['Media Agency','Event Agency','Creative production Agency','Social Agency','CRM Agency','PR Agency','Experimental Agency','Digital production Agency','Marketing Audits','Production cost (origination)','Production cost (adaption, transcation, repurposing)','SEO','Market Research (consumer insight)','Business Intelligence','MKT system development','Dealer Conference','Regional Meetings','Car cost running (depreciation)','System (platform) maintenance','DTS','Celebrity']:
                validation = False
    elif str(df_fcst[header_std[6]][row_idx]).upper() == 'WORKING' and df_fcst[header_std[7]][row_idx] not in ['Awareness', 'Consideration', 'Opinion']:
        if df_fcst[header_std[8]][row_idx] not in ['Traditional Media','Digital Media','Social Media','CRM','Call Center','Event','Public Relationship','POSM','Sponsorship']:
            if df_fcst[header_std[9]][row_idx] not in ['TV','Radio','Newspaper & Magazine','Cinema','OOH','Vertical','Vedio','News & Portal','SEM','Others','Social Media','CRM Campaign','Consumer inbound & outbound','Global Event','Test Drive','Autoshow','Product Display','Launch Campaign','Group Buy','Owner experience event','Seasonal Campaign','Cyber Campaign','Delivery Ceremony','Used car campaign','Plant Visit','MY change communication','New product communication','Event communication','POSM','Customer Magazine','Dealer opening support','Sponsorship']:
                validation = False
    return validation


def is_duplicate(f_name):
    issue_idx = -1
    issue_name = f_name.split('.')[0][-9:]
    issue_names = [issue_name, 'Act']
    for issue in df_fcst['Issue'][:]:
        if issue not in issue_names:
            issue_idx = df_fcst[df_fcst['Issue'] == issue].index.values
    return issue_idx


def to_float(col_idx, row_idx):
    if 20 <= col_idx <= 32:
        if str(df_fcst[header_std[col_idx]][row_idx]).replace(' ', '') == '' or str( df_fcst[header_std[col_idx]][row_idx]).replace('-', '') == '':
            df_fcst[header_std[col_idx]][row_idx] = float(0.0)
        else:
            df_fcst[header_std[col_idx]][row_idx] = float(df_fcst[header_std[col_idx]][row_idx])


def to_upper(row_idx, col, val):
    if str(df_fcst[header_std[col]][row_idx]).upper() == val.upper():
        df_fcst[header_std[col]][row_idx] = val


start = time.time()
header_std = ['Issue', 'Dept', 'Team', 'Carline', 'Lifecycle', 'Branding/NonBranding',
              'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type',
              'Activity', 'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order',
              'KPI Others', 'SMM Campaign Code (Y/N)', 'SC No.', ' SC Name',
              'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
              'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
header_err = ['File Name', 'Exception Type', 'Index', 'Issue', 'Dept', 'Team', 'Carline', 'Lifecycle',
              'Branding/NonBranding', 'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type', 'Activity',
              'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order', 'KPI Others', 'SMM Campaign Code (Y/N)',
              'SC No.', ' SC Name', 'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct',
              'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
err_header_err = ['File Name', 'Exception Type']
path = "D:/ASP - Erin/Raw Data/checking list.xlsx"
df_files = pd.read_excel(path, dtype=str)
file_name = df_files['File Path']
# file_name = ["D:/ASP - Erin/Raw Data/ASP Dbase Report 2019 FCST 1+11.xlsx"]
data_new = []
data_err = []
data_err_header = []
file_path_a = "D:/ASP - Erin/Raw Data/DirectoryA/"
file_path_b = "D:/ASP - Erin/Raw Data/DirectoryB/"
if not os.path.exists(file_path_a):
    os.makedirs(file_path_a)
if not os.path.exists(file_path_b):
    os.makedirs(file_path_b)
file_new = "D:/ASP - Erin/Raw Data/DirectoryA/summary.xlsx"
file_err = "D:/ASP - Erin/Raw Data/DirectoryA/error.xlsx"
file_header_err = "D:/ASP - Erin/Raw Data/DirectoryB/error_header.xlsx"
writer_new = pd.ExcelWriter(path=file_new, mode='w', engine='xlsxwriter')
writer_err = pd.ExcelWriter(path=file_err, mode='w', engine='openpyxl')
writer_err_header = pd.ExcelWriter(path=file_header_err, mode='w', engine='openpyxl')
for file_idx in range(len(file_name)):

    df_fcst = pd.read_excel(file_name[file_idx], dtype=str)
    header_result = df_fcst.columns[0:34] == header_std

    if not header_result.all():
        row_data_err = [file_name[file_idx], 'Header Exception']
        data_err_header.append(row_data_err)
        frame_header_err = pd.DataFrame(data_err_header, columns=err_header_err)
        frame_header_err.to_excel(writer_err_header, index=False, header=True)
        writer_err_header.save()
        writer_err_header.close()
    elif is_duplicate(file_name[file_idx]) != -1:
        row_data_err = [file_name[file_idx], 'Issue Duplicate', is_duplicate(file_name[file_idx]) + 2]
        data_err.append(row_data_err)
    else:
        for idx_row in range(len(df_fcst[header_std[0]])):
            row_data = []  # 临时存放符合的idx_row行的数据
            row_data_err = []  # 临时存放错误的idx_row行的数据
            # 判断idx_row的前两列是否为空
            if not is_null(idx_row):
                row_data_err.append(file_name[file_idx])
                row_data_err.append('Column Null Exception')
                row_data_err.append(idx_row + 2)
                for idx_col in range(len(header_std)):
                    row_data_err.append(df_fcst[header_std[idx_col]][idx_row])
                data_err.append(row_data_err)
            elif not calc_total_budget(idx_row):
                row_data_err.append(file_name[file_idx])
                row_data_err.append('Budget Unequal Exception')
                row_data_err.append(idx_row + 2)
                for idx_col in range(len(header_std)):
                    row_data_err.append(df_fcst[header_std[idx_col]][idx_row])
                data_err.append(row_data_err)
            elif not validate(idx_row):
                row_data_err.append(file_name[file_idx])
                row_data_err.append('Invalidation')
                row_data_err.append(idx_row + 2)
                for idx_col in range(len(header_std)):
                    row_data_err.append(df_fcst[header_std[idx_col]][idx_row])
                data_err.append(row_data_err)
            else:
                for idx_col in range(len(header_std)):
                    to_upper(idx_row,5,'Branding')
                    to_upper(idx_row,5,'NonBranding')
                    to_upper(idx_row,6,'Working')
                    to_upper(idx_row,6,'NonWorking')
                    to_float(idx_col,idx_row)
                    row_data.append(df_fcst[header_std[idx_col]][idx_row])
                data_new.append(row_data)
frame_err = pd.DataFrame(data_err, columns=header_err)
frame_err.to_excel(writer_err, index=False, header=True)
writer_err.save()
writer_err.close()
# 将正确数据写入excel
frame_new = pd.DataFrame(data_new, columns=header_std)
frame_new.to_excel(writer_new, index=False, header=True)
#Indicate workbook and worksheet for formatting
workbook = writer_new.book
worksheet = writer_new.sheets['Sheet1']
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
# Write the column headers with the defined format.
for col_num, value in enumerate(frame_new.columns.values):
    if col_num in [15,16,19,32]:
        worksheet.write(0, col_num, value, header_format_pqt)
    elif col_num in [17,18]:
        worksheet.write(0, col_num, value, header_format_rs)
    else:
        worksheet.write(0, col_num, value, header_format)
#Iterate through each column and set the width == the max length in that column. A padding length of 2 is also added.
for i, col in enumerate(frame_new.columns):
    # find length of column i
    column_len = frame_new[col].astype(str).str.len().mean()
    # Setting the length if the column header is larger than the max column value length
    column_len = max(column_len, len(col)) + 2
    # set the column length
    worksheet.set_column(i, i, column_len)

writer_new.save()
writer_new.close()
# print(data_new)
# print(data_err)
# print(data_err_header)
cost = time.time() - start
print(cost)
