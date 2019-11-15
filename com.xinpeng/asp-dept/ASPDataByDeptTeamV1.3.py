import time
import pandas as pd
import os


def mkdir(file_path):
    if not os.path.exists(file_path):
        os.makedirs(file_path)


def show_time(start_time):
    print("The program is running and it has been running for %.2fseconds" % (time.time() - start_time))


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
                   'Total Budget', 'Q1', 'Q2', 'Q3', 'Q4']:
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
        elif len(data_frame.columns) == 38:
            for row_idx in range(len(data)):
                for col_idx in range(38):
                    if 20 <= col_idx <= 32 or 34 <= col_idx <= 37:
                        worksheet.write(row_idx + 1, col_idx, data[row_idx][col_idx], budget_format)
    except IndexError:
        print("hello")
    writer.save()
    writer.close()


def write_to_excel(data, header, writer, err_col, red_col, gray_col):
    frame = pd.DataFrame(data, columns=header)
    frame.to_excel(writer, index=False, header=True)
    to_format(writer, frame, data, err_col, red_col, gray_col)


print("The program is starting ... ...")
start = time.time()
# 标准数据excel header
header_std = ['Issue', 'Dept', 'Team', 'Carline', 'Lifecycle', 'Branding/NonBranding',
              'Working/NonWorking', 'Sale funnel', 'Category', 'Activity type',
              'Activity', 'KPI Prospects', 'KPI Leads', 'KPI Inquiry', 'KPI Order',
              'KPI Others', 'SMM Campaign Code (Y/N)', 'SC NO.', ' SC Name',
              'Description', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug',
              'Sep', 'Oct', 'Nov', 'Dec', 'Total Budget', 'Stakeholder(CDSID)']
file_std = "W:/Finance Volvo/BIDataSource/Finance/ASP Database/Standard.xlsx"
df_folder_path = pd.read_excel(file_std, sheet_name='Folder Path', dtype=str)

folder_path = df_folder_path['Path'][0]
path = folder_path + "Raw Data/Parsing/"
try:
    os.listdir(path)
except FileNotFoundError:
    folder_path = folder_path + '/'
file_name_all = os.listdir(path)
file_name = []
for file_idx in range(len(file_name_all)):
    if file_name_all[file_idx].endswith('.xlsx'):
        file_name.append(path + file_name_all[file_idx])

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
# 创建存放excel的文件夹
path_result_parse = folder_path + "Result/Parsing"
mkdir(path_result_parse)
show_time(start)
# 几个生成的excel路径及文件名
file_MKT = folder_path + "Result/Parsing/MKT.xlsx"
file_MKT_Launch = folder_path + "Result/Parsing/MKT Launch.xlsx"
file_APAC_CC = folder_path + "Result/Parsing/APAC CC.xlsx"
file_Customer_Service = folder_path + "Result/Parsing/Customer Service.xlsx"
file_DTS = folder_path + "Result/Parsing/DTS.xlsx"
file_MI = folder_path + "Result/Parsing/MI.xlsx"
file_Sales_MKT1 = folder_path + "Result/Parsing/Sales MKT_DMKT.xlsx"
file_Sales_MKT2 = folder_path + "Result/Parsing/Sales MKT_HK.xlsx"
file_Sales_MKT3 = folder_path + "Result/Parsing/Sales MKT_Internal Fleet Car.xlsx"
file_Sales_MKT4 = folder_path + "Result/Parsing/Sales MKT_National Sales.xlsx"
file_Sales_MKT5 = folder_path + "Result/Parsing/Sales MKT_NBD Team.xlsx"
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
# 从文件列表遍历读取文件
for file_idx in range(len(file_name)):
    df_asp = pd.read_excel(file_name[file_idx], dtype=str)
    show_time(start)
    for idx_row in range(len(df_asp[header_std[0]])):
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
            if str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKT':
                for idx_col in range(len(header_std)):
                    row_data_MKT.append(df_asp[header_std[idx_col]][idx_row])
                data_MKT.append(row_data_MKT)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKTLAUNCH':
                for idx_col in range(len(header_std)):
                    row_data_MKT_Launch.append(df_asp[header_std[idx_col]][idx_row])
                data_MKT_Launch.append(row_data_MKT_Launch)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'APACCC':
                for idx_col in range(len(header_std)):
                    row_data_APAC_CC.append(df_asp[header_std[idx_col]][idx_row])
                data_APAC_CC.append(row_data_APAC_CC)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'CUSTOMERSERVICE':
                for idx_col in range(len(header_std)):
                    row_data_Customer_Service.append(df_asp[header_std[idx_col]][idx_row])
                data_Customer_Service.append(row_data_Customer_Service)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'DTS':
                for idx_col in range(len(header_std)):
                    row_data_DTS.append(df_asp[header_std[idx_col]][idx_row])
                data_DTS.append(row_data_DTS)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MI':
                for idx_col in range(len(header_std)):
                    row_data_MI.append(df_asp[header_std[idx_col]][idx_row])
                data_MI.append(row_data_MI)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] \
                    and str(df_asp[header_std[2]][idx_row]).replace(' ', '').upper() \
                    in ['DMKTCENTRALTEAM', 'EXHIBITION', 'EASTREGIONTEAM', 'NORTHREGIONTEAM', 'SOUTHREGIONTEAM',
                        'WESTREGIONTEAM', 'ZHEJIANGREGIONTEAM']:
                for idx_col in range(len(header_std)):
                    row_data_Sales_MKT1.append(df_asp[header_std[idx_col]][idx_row])
                data_Sales_MKT1.append(row_data_Sales_MKT1)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                    df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['HK']:
                for idx_col in range(len(header_std)):
                    row_data_Sales_MKT2.append(df_asp[header_std[idx_col]][idx_row])
                data_Sales_MKT2.append(row_data_Sales_MKT2)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                    df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['INTERNALFLEETCAR']:
                for idx_col in range(len(header_std)):
                    row_data_Sales_MKT3.append(df_asp[header_std[idx_col]][idx_row])
                data_Sales_MKT3.append(row_data_Sales_MKT3)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                    df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['NATIONALSALES']:
                for idx_col in range(len(header_std)):
                    row_data_Sales_MKT4.append(df_asp[header_std[idx_col]][idx_row])
                data_Sales_MKT4.append(row_data_Sales_MKT4)
            elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] \
                    and str(df_asp[header_std[2]][idx_row]).replace(' ', '').upper() \
                    in ['FLEETTEAM', 'USEDCARTEAM', 'FINANCIALSERVICETEAM']:
                for idx_col in range(len(header_std)):
                    row_data_Sales_MKT5.append(df_asp[header_std[idx_col]][idx_row])
                data_Sales_MKT5.append(row_data_Sales_MKT5)

write_to_excel(data_MKT, header_std, writer_MKT, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_MKT_Launch, header_std, writer_MKT_Launch, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_APAC_CC, header_std, writer_APAC_CC, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Customer_Service, header_std, writer_Customer_Service, [], [15, 16, 19, 32], [17, 18])
show_time(start)
write_to_excel(data_DTS, header_std, writer_DTS, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_MI, header_std, writer_MI, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT1, header_std, writer_Sales_MKT1, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT2, header_std, writer_Sales_MKT2, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT3, header_std, writer_Sales_MKT3, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT4, header_std, writer_Sales_MKT4, [], [15, 16, 19, 32], [17, 18])
write_to_excel(data_Sales_MKT5, header_std, writer_Sales_MKT5, [], [15, 16, 19, 32], [17, 18])
os.system('cls')
print("The program has ended. The program runs for a total of %.2fseconds" % (time.time() - start))
os.system("pause")