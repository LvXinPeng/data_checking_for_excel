import time
import pandas as pd
import os


def mkdir(file_path):
    if not os.path.exists(file_path):
        os.makedirs(file_path)


def to_float(col_idx, row_idx):
    if 20 <= col_idx <= 32:
        df_summary[header_std[col_idx]][row_idx] = float(df_summary[header_std[col_idx]][row_idx])


def to_format(writer, data_frame, data, red_col, gray_col):
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


start = time.time()
# 标准数据excel header
header_std = ['Issue', 'Dept', 'Team', 'Carline', 'Lifecycle', 'Branding/NonBranding',
              'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type',
              'Activity', 'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order',
              'KPI Others', 'SMM Campaign Code (Y/N)', 'SC No.', ' SC Name',
              'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
              'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
path = "D:/ASP - Erin/Dept-Team/"
mkdir(path)
# 读入待处理文件
file_summary = "D:/ASP - Erin/Data&Log/Summary.xlsx"
df_summary = pd.read_excel(file_summary, dtype=str)

file_MKT = "D:/ASP - Erin/Dept-Team/MKT.xlsx"
file_MKT_Launch = "D:/ASP - Erin/Dept-Team/MKT Launch.xlsx"
file_APAC_CC = "D:/ASP - Erin/Dept-Team/APAC CC.xlsx"
file_Customer_Service = "D:/ASP - Erin/Dept-Team/Customer Service.xlsx"
file_DTS = "D:/ASP - Erin/Dept-Team/DTS.xlsx"
file_MI = "D:/ASP - Erin/Dept-Team/MI.xlsx"
file_Sales_MKT1 = "D:/ASP - Erin/Dept-Team/Sales MKT_DMKT.xlsx"
file_Sales_MKT2 = "D:/ASP - Erin/Dept-Team/Sales MKT_HK.xlsx"
file_Sales_MKT3 = "D:/ASP - Erin/Dept-Team/Sales MKT_Internal Fleet Car.xlsx"
file_Sales_MKT4 = "D:/ASP - Erin/Dept-Team/Sales MKT_National Sales.xlsx"
file_Sales_MKT5 = "D:/ASP - Erin/Dept-Team/Sales MKT_NBD Team.xlsx"

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
            to_float(idx_col, idx_row)
            row_data_MKT.append(df_summary[header_std[idx_col]][idx_row])
        data_MKT.append(row_data_MKT)
    elif str(df_summary[header_std[1]][idx_row]) == 'MKT Launch':
        for idx_col in range(len(header_std)):
            to_float(idx_col, idx_row)
            row_data_MKT_Launch.append(df_summary[header_std[idx_col]][idx_row])
        data_MKT_Launch.append(row_data_MKT_Launch)
    elif str(df_summary[header_std[1]][idx_row]) == 'APAC CC':
        for idx_col in range(len(header_std)):
            to_float(idx_col, idx_row)
            row_data_APAC_CC.append(df_summary[header_std[idx_col]][idx_row])
        data_APAC_CC.append(row_data_APAC_CC)
    elif str(df_summary[header_std[1]][idx_row]) == 'Customer Service':
        for idx_col in range(len(header_std)):
            to_float(idx_col, idx_row)
            row_data_Customer_Service.append(df_summary[header_std[idx_col]][idx_row])
        data_Customer_Service.append(row_data_Customer_Service)
    elif str(df_summary[header_std[1]][idx_row]) == 'DTS':
        for idx_col in range(len(header_std)):
            to_float(idx_col, idx_row)
            row_data_DTS.append(df_summary[header_std[idx_col]][idx_row])
        data_DTS.append(row_data_DTS)
    elif str(df_summary[header_std[1]][idx_row]) == 'MI':
        for idx_col in range(len(header_std)):
            to_float(idx_col, idx_row)
            row_data_MI.append(df_summary[header_std[idx_col]][idx_row])
        data_MI.append(row_data_MI)
    elif str(df_summary[header_std[1]][idx_row]) == 'Sales MKT' and str(df_summary[header_std[2]][idx_row]) in [
        'DMKT Central Team', 'Exhibition', 'East Region Team', 'North Region Team', 'South Region Team',
        'West Region Team', 'ZheJiang Region Team']:
        for idx_col in range(len(header_std)):
            to_float(idx_col, idx_row)
            row_data_Sales_MKT1.append(df_summary[header_std[idx_col]][idx_row])
        data_Sales_MKT1.append(row_data_Sales_MKT1)
    elif str(df_summary[header_std[1]][idx_row]) == 'Sales MKT' and str(df_summary[header_std[2]][idx_row]) in ['HK']:
        for idx_col in range(len(header_std)):
            to_float(idx_col, idx_row)
            row_data_Sales_MKT2.append(df_summary[header_std[idx_col]][idx_row])
        data_Sales_MKT2.append(row_data_Sales_MKT2)
    elif str(df_summary[header_std[1]][idx_row]) == 'Sales MKT' and str(df_summary[header_std[2]][idx_row]) in [
        'Internal Fleet Car']:
        for idx_col in range(len(header_std)):
            to_float(idx_col, idx_row)
            row_data_Sales_MKT3.append(df_summary[header_std[idx_col]][idx_row])
        data_Sales_MKT3.append(row_data_Sales_MKT3)
    elif str(df_summary[header_std[1]][idx_row]) == 'Sales MKT' and str(df_summary[header_std[2]][idx_row]) in [
        'National Sales']:
        for idx_col in range(len(header_std)):
            to_float(idx_col, idx_row)
            row_data_Sales_MKT4.append(df_summary[header_std[idx_col]][idx_row])
        data_Sales_MKT4.append(row_data_Sales_MKT4)
    elif str(df_summary[header_std[1]][idx_row]) == 'Sales MKT' and str(df_summary[header_std[2]][idx_row]) in [
        'NBD Team']:
        for idx_col in range(len(header_std)):
            to_float(idx_col, idx_row)
            row_data_Sales_MKT5.append(df_summary[header_std[idx_col]][idx_row])
        data_Sales_MKT5.append(row_data_Sales_MKT5)

# 将标准数据写入Summary excel
frame_new1 = pd.DataFrame(data_MKT, columns=header_std)
frame_new1.to_excel(writer_MKT, index=False, header=True)
to_format(writer_MKT, frame_new1, data_MKT, [15, 16, 19, 32], [17, 18])

frame_new2 = pd.DataFrame(data_MKT_Launch, columns=header_std)
frame_new2.to_excel(writer_MKT_Launch, index=False, header=True)
to_format(writer_MKT_Launch, frame_new2, data_MKT_Launch, [15, 16, 19, 32], [17, 18])

frame_new3 = pd.DataFrame(data_APAC_CC, columns=header_std)
frame_new3.to_excel(writer_APAC_CC, index=False, header=True)
to_format(writer_APAC_CC, frame_new3, data_APAC_CC, [15, 16, 19, 32], [17, 18])

frame_new4 = pd.DataFrame(data_Customer_Service, columns=header_std)
frame_new4.to_excel(writer_Customer_Service, index=False, header=True)
to_format(writer_Customer_Service, frame_new4, data_Customer_Service, [15, 16, 19, 32], [17, 18])

frame_new5 = pd.DataFrame(data_DTS, columns=header_std)
frame_new5.to_excel(writer_DTS, index=False, header=True)
to_format(writer_DTS, frame_new5, data_DTS, [15, 16, 19, 32], [17, 18])

frame_new6 = pd.DataFrame(data_MI, columns=header_std)
frame_new6.to_excel(writer_MI, index=False, header=True)
to_format(writer_MI, frame_new6, data_MI, [15, 16, 19, 32], [17, 18])

frame_new7 = pd.DataFrame(data_Sales_MKT1, columns=header_std)
frame_new7.to_excel(writer_Sales_MKT1, index=False, header=True)
to_format(writer_Sales_MKT1, frame_new7, data_Sales_MKT1, [15, 16, 19, 32], [17, 18])

frame_new8 = pd.DataFrame(data_Sales_MKT2, columns=header_std)
frame_new8.to_excel(writer_Sales_MKT2, index=False, header=True)
to_format(writer_Sales_MKT2, frame_new8, data_Sales_MKT2, [15, 16, 19, 32], [17, 18])

frame_new9 = pd.DataFrame(data_Sales_MKT3, columns=header_std)
frame_new9.to_excel(writer_Sales_MKT3, index=False, header=True)
to_format(writer_Sales_MKT3, frame_new9, data_Sales_MKT3, [15, 16, 19, 32], [17, 18])

frame_new10 = pd.DataFrame(data_Sales_MKT4, columns=header_std)
frame_new10.to_excel(writer_Sales_MKT4, index=False, header=True)
to_format(writer_Sales_MKT4, frame_new10, data_Sales_MKT4, [15, 16, 19, 32], [17, 18])

frame_new11 = pd.DataFrame(data_Sales_MKT5, columns=header_std)
frame_new11.to_excel(writer_Sales_MKT5, index=False, header=True)
to_format(writer_Sales_MKT5, frame_new11, data_Sales_MKT5, [15, 16, 19, 32], [17, 18])
# 计算程序耗时
cost = time.time() - start
print(cost)
