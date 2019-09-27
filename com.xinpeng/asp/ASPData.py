import pandas as pd


def calc_total_budget(row_idx):
    temp = []
    is_eq = False
    if df_fcst[header_std[20]][row_idx] == '' or df_fcst[header_std[21]][row_idx] == '' or df_fcst[header_std[22]][row_idx] == '' or df_fcst[header_std[23]][row_idx] == '' or df_fcst[header_std[24]][row_idx] == '' or df_fcst[header_std[25]][row_idx] == '' or df_fcst[header_std[26]][row_idx] == '' or df_fcst[header_std[27]][row_idx] == '' or df_fcst[header_std[28]][row_idx] == '' or df_fcst[header_std[29]][row_idx] == '' or df_fcst[header_std[30]][row_idx] == '' or df_fcst[header_std[31]][row_idx] == '' or df_fcst[header_std[32]][row_idx] == '':
        return is_eq
    else:
        for col_idx in range(12):
            if str(df_fcst[header_std[col_idx + 20]][row_idx]).replace(' ','') == '':
                temp.append(0)
            else:
                temp.append(df_fcst[header_std[col_idx + 20]][row_idx])
        total_budget = sum(list(map(float, temp)))
        is_eq = total_budget == float(df_fcst[header_std[32]][row_idx])
    return is_eq


def validate(row_idx):
    validation = True
    if df_fcst[header_std[6]][row_idx] == 'NonWorking' and df_fcst[header_std[7]][row_idx] != "0":
        validation = False
        # return validation
    elif df_fcst[header_std[6]][row_idx] == 'Working' and df_fcst[header_std[7]][row_idx] not in ['Awareness','Consideration','Opinion']:
        validation = False
        # return validation
    return validation


def is_duplicate(f_name):
    issue_idx = -1
    issue_name = f_name.split('.')[0][-9:]
    issue_names = [issue_name, 'Act']
    for issue in df_fcst['Issue'][:]:
        if issue not in issue_names:
            issue_idx = df_fcst[df_fcst['Issue'] == issue].index.values
    return issue_idx


header_std = ['Issue', 'Dept', 'Team', 'Carline', 'Lifecycle', 'Branding/NonBranding',
              'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type',
              'Activity', 'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order',
              'KPI Others', 'SMM Campaign Code (Y/N)', 'SC No.', ' SC Name',
              'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
              'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
header_err = ['File Name', 'Exception Type', 'Index']

path_for_new = "D:/ASP - Erin/Raw Data/"
file_name = 'D:/ASP - Erin/Raw Data/ASP Dbase Report 2019 FCST 1+11.xlsx'
file_new = file_name.split('.xlsx')[0] + "_new.xlsx"
file_err = 'D:/ASP - Erin/Raw Data/error.xlsx'

df_fcst = pd.read_excel(file_name, dtype=str)
writer_new = pd.ExcelWriter(path=file_new, mode='w', engine='openpyxl')
writer_err = pd.ExcelWriter(path=file_err, mode='w', engine='openpyxl')
header_result = df_fcst.columns[0:34] == header_std

data_new = []
data_err = []

if not header_result.all():
    row_data = [file_name, 'Header Exception', 'All']
    data_err.append(row_data)
    frame = pd.DataFrame(data_err, columns=header_err)
    frame.to_excel(writer_err, index=False, header=True)
    writer_err.save()
    writer_err.close()
    print('Finished')
elif is_duplicate(file_name) != -1:
    row_data_err = [file_name, 'Issue Duplicate', is_duplicate(file_name) + 2]
    data_err.append(row_data_err)
    frame = pd.DataFrame(data_err, columns=header_err)
    frame.to_excel(writer_err, index=False, header=True)
    writer_err.save()
    writer_err.close()
    print('Finished')
else:
    for idx_row in range(len(df_fcst[header_std[0]])):
        row_data = []  # 临时存放符合的idx_row行的数据
        row_data_err = []  # 临时存放错误的idx_row行的数据
        # 判断idx_row的前两列是否为空
        if df_fcst[header_std[0]][idx_row] == '' or df_fcst[header_std[1]][idx_row] == '':
            row_data_err.append(file_name)
            row_data_err.append('Column Null Exception')
            row_data_err.append(idx_row + 2)
            data_err.append(row_data_err)
        elif not calc_total_budget(idx_row):
            row_data_err.append(file_name)
            row_data_err.append('Budget Unequal Exception')
            row_data_err.append(idx_row + 2)
            data_err.append(row_data_err)
        elif validate(idx_row) == False:
            row_data_err.append(file_name)
            row_data_err.append('Invalidation')
            row_data_err.append(idx_row + 2)
            data_err.append(row_data_err)
        else:
            for idx_col in range(len(header_std)):
                row_data.append(df_fcst[header_std[idx_col]][idx_row])
            data_new.append(row_data)
    # 将异常数据写入error Excel
    frame_err = pd.DataFrame(data_err, columns=header_err)
    frame_err.to_excel(writer_err, index=False, header=True)
    writer_err.save()
    writer_err.close()
    # 将正确数据写入excel
    frame_new = pd.DataFrame(data_new, columns=header_std)
    frame_new.to_excel(writer_new, index=False, header=True)
    writer_new.save()
    writer_new.close()
    print(data_new)
    print('Finished')
