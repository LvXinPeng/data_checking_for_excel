import datetime
import time
import pandas as pd
import os


def mkdir(file_path):
    if not os.path.exists(file_path):
        os.makedirs(file_path)


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


start = time.time()
# 标准数据excel header
header_std = ['Issue', 'Dept', 'Team', 'Carline', 'Lifecycle', 'Branding/NonBranding',
              'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type',
              'Activity', 'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order',
              'KPI Others', 'SMM Campaign Code (Y/N)', 'SC No.', ' SC Name',
              'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
              'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
path2 = "D:/ASP - Erin/Actual/Dept-Team/"
mkdir(path2)
# 读入待处理文件
today = str(datetime.date.today()).replace('-', '')
file_summary2 = "D:/ASP - Erin/Actual/Data&Log/Act_Summary_" + today + ".xlsx"
df_summary2 = pd.read_excel(file_summary2, dtype=str)

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
