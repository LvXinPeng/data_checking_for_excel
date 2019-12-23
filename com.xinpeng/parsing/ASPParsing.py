try:
    import time
    import pandas as pd
    import os


    def mkdir(file_path):
        if not os.path.exists(file_path):
            os.makedirs(file_path)


    def show_time(start_time):
        print("The program is running and it has been running for %.2fseconds" % (time.time() - start_time))


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
        if 11 <= col_idx <= 14:
            if is_number(str(df[header[col_idx]][row_idx])):
                try:
                    df[header[col_idx]][row_idx] = int(df[header[col_idx]][row_idx])
                except ValueError:
                    df[header[col_idx]][row_idx] = float(df[header[col_idx]][row_idx])


    def is_actual(df, header, row_idx, actual_name):
        flag = False
        if str(df[header[0]][row_idx]).replace(' ', '').upper() in actual_name:
            flag = True
        return flag


    def is_forecast(df, header, row_idx, issue_name):
        flag = False
        if str(df[header[0]][row_idx]).replace(' ', '').upper() in issue_name:
            flag = True
        return flag


    def is_budget(df, header, row_idx, budget_name):
        flag = False
        if str(df[header[0]][row_idx]).replace(' ', '').upper() in budget_name:
            flag = True
        return flag


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

    df_issue = pd.read_excel(file_std, sheet_name='Issue Standard', dtype=str)
    issue_fcst_list = []
    issue_fcst_parsing_list = []
    for issue_fcst in df_issue['Issue FCST']:
        if str(issue_fcst) != 'nan':
            issue_fcst_parsing_list.append(issue_fcst)
            issue_fcst_list.append(issue_fcst.replace(' ', '').upper())
    issue_actual_list = []
    issue_actual_parsing_list = []
    for issue_actual in df_issue['Issue Actual']:
        if str(issue_actual) != 'nan':
            issue_actual_parsing_list.append(issue_actual)
            issue_actual_list.append(issue_actual.replace(' ', '').upper())
    issue_budget_list = []
    for issue_budget in df_issue['Issue Budget']:
        if str(issue_budget) != 'nan':
            issue_budget_list.append(issue_budget.replace(' ', '').upper())
    fcst_parsing_name = min(issue_fcst_parsing_list)
    actual_parsing_name = min(issue_actual_parsing_list)
    bgt_parsing_name = min(issue_budget_list)

    folder_path = df_folder_path['Path'][0]
    path = folder_path + "Raw Data/Parsing/"
    try:
        os.listdir(path)
    except FileNotFoundError:
        folder_path = folder_path + '/'
        path = folder_path + "Raw Data/Parsing/"
    file_name_all = os.listdir(path)
    file_name = []
    for file_idx in range(len(file_name_all)):
        if file_name_all[file_idx].endswith('.xlsx'):
            file_name.append(path + file_name_all[file_idx])

    data_MKT_FCST = []  # 存储所有的标准数据
    data_MKT_Launch_FCST = []
    data_APAC_CC_FCST = []
    data_Customer_Service_FCST = []
    data_DTS_FCST = []
    data_MI_FCST = []
    data_Sales_MKT1_FCST = []
    data_Sales_MKT2_FCST = []
    data_Sales_MKT3_FCST = []
    data_Sales_MKT4_FCST = []
    data_Sales_MKT5_FCST = []

    data_MKT_ACT = []  # 存储所有的标准数据
    data_MKT_Launch_ACT = []
    data_APAC_CC_ACT = []
    data_Customer_Service_ACT = []
    data_DTS_ACT = []
    data_MI_ACT = []
    data_Sales_MKT1_ACT = []
    data_Sales_MKT2_ACT = []
    data_Sales_MKT3_ACT = []
    data_Sales_MKT4_ACT = []
    data_Sales_MKT5_ACT = []
    # 创建存放excel的文件夹
    path_result_parse = folder_path + "Result/Parsing"
    mkdir(path_result_parse)
    show_time(start)
    # 几个生成的excel路径及文件名
    file_MKT_FCST = folder_path + "Result/Parsing/" + fcst_parsing_name + "_MKT.xlsx"
    file_MKT_Launch_FCST = folder_path + "Result/Parsing/" + fcst_parsing_name + "_MKT Launch.xlsx"
    file_APAC_CC_FCST = folder_path + "Result/Parsing/" + fcst_parsing_name + "_APAC CC.xlsx"
    file_Customer_Service_FCST = folder_path + "Result/Parsing/" + fcst_parsing_name + "_Customer Service.xlsx"
    file_DTS_FCST = folder_path + "Result/Parsing/" + fcst_parsing_name + "_DTS.xlsx"
    file_MI_FCST = folder_path + "Result/Parsing/" + fcst_parsing_name + "_MI.xlsx"
    file_Sales_MKT1_FCST = folder_path + "Result/Parsing/" + fcst_parsing_name + "_Sales MKT_DMKT.xlsx"
    file_Sales_MKT2_FCST = folder_path + "Result/Parsing/" + fcst_parsing_name + "_Sales MKT_HK.xlsx"
    file_Sales_MKT3_FCST = folder_path + "Result/Parsing/" + fcst_parsing_name + "_Sales MKT_Internal Fleet Car.xlsx"
    file_Sales_MKT4_FCST = folder_path + "Result/Parsing/" + fcst_parsing_name + "_Sales MKT_National Sales.xlsx"
    file_Sales_MKT5_FCST = folder_path + "Result/Parsing/" + fcst_parsing_name + "_Sales MKT_NBD Team.xlsx"

    file_MKT_ACT = folder_path + "Result/Parsing/" + actual_parsing_name + "_MKT.xlsx"
    file_MKT_Launch_ACT = folder_path + "Result/Parsing/" + actual_parsing_name + "_MKT Launch.xlsx"
    file_APAC_CC_ACT = folder_path + "Result/Parsing/" + actual_parsing_name + "_APAC CC.xlsx"
    file_Customer_Service_ACT = folder_path + "Result/Parsing/" + actual_parsing_name + "_Customer Service.xlsx"
    file_DTS_ACT = folder_path + "Result/Parsing/" + actual_parsing_name + "_DTS.xlsx"
    file_MI_ACT = folder_path + "Result/Parsing/" + actual_parsing_name + "_MI.xlsx"
    file_Sales_MKT1_ACT = folder_path + "Result/Parsing/" + actual_parsing_name + "_Sales MKT_DMKT.xlsx"
    file_Sales_MKT2_ACT = folder_path + "Result/Parsing/" + actual_parsing_name + "_Sales MKT_HK.xlsx"
    file_Sales_MKT3_ACT = folder_path + "Result/Parsing/" + actual_parsing_name + "_Sales MKT_Internal Fleet Car.xlsx"
    file_Sales_MKT4_ACT = folder_path + "Result/Parsing/" + actual_parsing_name + "_Sales MKT_National Sales.xlsx"
    file_Sales_MKT5_ACT = folder_path + "Result/Parsing/" + actual_parsing_name + "_Sales MKT_NBD Team.xlsx"
    # 创建几个excel的Writer引擎
    writer_MKT_FCST = pd.ExcelWriter(path=file_MKT_FCST, mode='w', engine='xlsxwriter')
    writer_MKT_Launch_FCST = pd.ExcelWriter(path=file_MKT_Launch_FCST, mode='w', engine='xlsxwriter')
    writer_APAC_CC_FCST = pd.ExcelWriter(path=file_APAC_CC_FCST, mode='w', engine='xlsxwriter')
    writer_Customer_Service_FCST = pd.ExcelWriter(path=file_Customer_Service_FCST, mode='w', engine='xlsxwriter')
    writer_DTS_FCST = pd.ExcelWriter(path=file_DTS_FCST, mode='w', engine='xlsxwriter')
    writer_MI_FCST = pd.ExcelWriter(path=file_MI_FCST, mode='w', engine='xlsxwriter')
    writer_Sales_MKT1_FCST = pd.ExcelWriter(path=file_Sales_MKT1_FCST, mode='w', engine='xlsxwriter')
    writer_Sales_MKT2_FCST = pd.ExcelWriter(path=file_Sales_MKT2_FCST, mode='w', engine='xlsxwriter')
    writer_Sales_MKT3_FCST = pd.ExcelWriter(path=file_Sales_MKT3_FCST, mode='w', engine='xlsxwriter')
    writer_Sales_MKT4_FCST = pd.ExcelWriter(path=file_Sales_MKT4_FCST, mode='w', engine='xlsxwriter')
    writer_Sales_MKT5_FCST = pd.ExcelWriter(path=file_Sales_MKT5_FCST, mode='w', engine='xlsxwriter')

    writer_MKT_ACT = pd.ExcelWriter(path=file_MKT_ACT, mode='w', engine='xlsxwriter')
    writer_MKT_Launch_ACT = pd.ExcelWriter(path=file_MKT_Launch_ACT, mode='w', engine='xlsxwriter')
    writer_APAC_CC_ACT = pd.ExcelWriter(path=file_APAC_CC_ACT, mode='w', engine='xlsxwriter')
    writer_Customer_Service_ACT = pd.ExcelWriter(path=file_Customer_Service_ACT, mode='w', engine='xlsxwriter')
    writer_DTS_ACT = pd.ExcelWriter(path=file_DTS_ACT, mode='w', engine='xlsxwriter')
    writer_MI_ACT = pd.ExcelWriter(path=file_MI_ACT, mode='w', engine='xlsxwriter')
    writer_Sales_MKT1_ACT = pd.ExcelWriter(path=file_Sales_MKT1_ACT, mode='w', engine='xlsxwriter')
    writer_Sales_MKT2_ACT = pd.ExcelWriter(path=file_Sales_MKT2_ACT, mode='w', engine='xlsxwriter')
    writer_Sales_MKT3_ACT = pd.ExcelWriter(path=file_Sales_MKT3_ACT, mode='w', engine='xlsxwriter')
    writer_Sales_MKT4_ACT = pd.ExcelWriter(path=file_Sales_MKT4_ACT, mode='w', engine='xlsxwriter')
    writer_Sales_MKT5_ACT = pd.ExcelWriter(path=file_Sales_MKT5_ACT, mode='w', engine='xlsxwriter')
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
            for idx_col in range(len(header_std)):
                to_float(df_asp, header_std, idx_col, idx_row)
                to_int(df_asp, header_std, idx_col, idx_row)
            if is_forecast(df_asp, header_std, idx_row, issue_fcst_list):
                if str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKT':
                    for idx_col in range(len(header_std)):
                        row_data_MKT.append(df_asp[header_std[idx_col]][idx_row])
                    data_MKT_FCST.append(row_data_MKT)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKTLAUNCH':
                    for idx_col in range(len(header_std)):
                        row_data_MKT_Launch.append(df_asp[header_std[idx_col]][idx_row])
                    data_MKT_Launch_FCST.append(row_data_MKT_Launch)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'APACCC':
                    for idx_col in range(len(header_std)):
                        row_data_APAC_CC.append(df_asp[header_std[idx_col]][idx_row])
                    data_APAC_CC_FCST.append(row_data_APAC_CC)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'CUSTOMERSERVICE':
                    for idx_col in range(len(header_std)):
                        row_data_Customer_Service.append(df_asp[header_std[idx_col]][idx_row])
                    data_Customer_Service_FCST.append(row_data_Customer_Service)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'DTS':
                    for idx_col in range(len(header_std)):
                        row_data_DTS.append(df_asp[header_std[idx_col]][idx_row])
                    data_DTS_FCST.append(row_data_DTS)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MI':
                    for idx_col in range(len(header_std)):
                        row_data_MI.append(df_asp[header_std[idx_col]][idx_row])
                    data_MI_FCST.append(row_data_MI)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] \
                        and str(df_asp[header_std[2]][idx_row]).replace(' ', '').upper() \
                        in ['DMKTCENTRALTEAM', 'EXHIBITION', 'EASTREGIONTEAM', 'NORTHREGIONTEAM', 'SOUTHREGIONTEAM',
                            'WESTREGIONTEAM', 'ZHEJIANGREGIONTEAM']:
                    for idx_col in range(len(header_std)):
                        row_data_Sales_MKT1.append(df_asp[header_std[idx_col]][idx_row])
                    data_Sales_MKT1_FCST.append(row_data_Sales_MKT1)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                        df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['HK']:
                    for idx_col in range(len(header_std)):
                        row_data_Sales_MKT2.append(df_asp[header_std[idx_col]][idx_row])
                    data_Sales_MKT2_FCST.append(row_data_Sales_MKT2)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                        df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['INTERNALFLEETCAR']:
                    for idx_col in range(len(header_std)):
                        row_data_Sales_MKT3.append(df_asp[header_std[idx_col]][idx_row])
                    data_Sales_MKT3_FCST.append(row_data_Sales_MKT3)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                        df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['NATIONALSALES']:
                    for idx_col in range(len(header_std)):
                        row_data_Sales_MKT4.append(df_asp[header_std[idx_col]][idx_row])
                    data_Sales_MKT4_FCST.append(row_data_Sales_MKT4)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] \
                        and str(df_asp[header_std[2]][idx_row]).replace(' ', '').upper() \
                        in ['FLEETTEAM', 'USEDCARTEAM', 'FINANCIALSERVICETEAM']:
                    for idx_col in range(len(header_std)):
                        row_data_Sales_MKT5.append(df_asp[header_std[idx_col]][idx_row])
                    data_Sales_MKT5_FCST.append(row_data_Sales_MKT5)
            elif is_actual(df_asp, header_std, idx_row, issue_actual_list):
                if str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKT':
                    for idx_col in range(len(header_std)):
                        row_data_MKT.append(df_asp[header_std[idx_col]][idx_row])
                    data_MKT_ACT.append(row_data_MKT)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MKTLAUNCH':
                    for idx_col in range(len(header_std)):
                        row_data_MKT_Launch.append(df_asp[header_std[idx_col]][idx_row])
                    data_MKT_Launch_ACT.append(row_data_MKT_Launch)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'APACCC':
                    for idx_col in range(len(header_std)):
                        row_data_APAC_CC.append(df_asp[header_std[idx_col]][idx_row])
                    data_APAC_CC_ACT.append(row_data_APAC_CC)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'CUSTOMERSERVICE':
                    for idx_col in range(len(header_std)):
                        row_data_Customer_Service.append(df_asp[header_std[idx_col]][idx_row])
                    data_Customer_Service_ACT.append(row_data_Customer_Service)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'DTS':
                    for idx_col in range(len(header_std)):
                        row_data_DTS.append(df_asp[header_std[idx_col]][idx_row])
                    data_DTS_ACT.append(row_data_DTS)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() == 'MI':
                    for idx_col in range(len(header_std)):
                        row_data_MI.append(df_asp[header_std[idx_col]][idx_row])
                    data_MI_ACT.append(row_data_MI)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] \
                        and str(df_asp[header_std[2]][idx_row]).replace(' ', '').upper() \
                        in ['DMKTCENTRALTEAM', 'EXHIBITION', 'EASTREGIONTEAM', 'NORTHREGIONTEAM', 'SOUTHREGIONTEAM',
                            'WESTREGIONTEAM', 'ZHEJIANGREGIONTEAM']:
                    for idx_col in range(len(header_std)):
                        row_data_Sales_MKT1.append(df_asp[header_std[idx_col]][idx_row])
                    data_Sales_MKT1_ACT.append(row_data_Sales_MKT1)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                        df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['HK']:
                    for idx_col in range(len(header_std)):
                        row_data_Sales_MKT2.append(df_asp[header_std[idx_col]][idx_row])
                    data_Sales_MKT2_ACT.append(row_data_Sales_MKT2)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                        df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['INTERNALFLEETCAR']:
                    for idx_col in range(len(header_std)):
                        row_data_Sales_MKT3.append(df_asp[header_std[idx_col]][idx_row])
                    data_Sales_MKT3_ACT.append(row_data_Sales_MKT3)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] and str(
                        df_asp[header_std[2]][idx_row]).replace(' ', '').upper() in ['NATIONALSALES']:
                    for idx_col in range(len(header_std)):
                        row_data_Sales_MKT4.append(df_asp[header_std[idx_col]][idx_row])
                    data_Sales_MKT4_ACT.append(row_data_Sales_MKT4)
                elif str(df_asp[header_std[1]][idx_row]).replace(' ', '').upper() in ['SALESMKT'] \
                        and str(df_asp[header_std[2]][idx_row]).replace(' ', '').upper() \
                        in ['FLEETTEAM', 'USEDCARTEAM', 'FINANCIALSERVICETEAM']:
                    for idx_col in range(len(header_std)):
                        row_data_Sales_MKT5.append(df_asp[header_std[idx_col]][idx_row])
                    data_Sales_MKT5_ACT.append(row_data_Sales_MKT5)

    write_to_excel(data_MKT_FCST, header_std, writer_MKT_FCST, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_MKT_Launch_FCST, header_std, writer_MKT_Launch_FCST, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_APAC_CC_FCST, header_std, writer_APAC_CC_FCST, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Customer_Service_FCST, header_std, writer_Customer_Service_FCST, [], [15, 16, 19, 32], [17, 18])
    show_time(start)
    write_to_excel(data_DTS_FCST, header_std, writer_DTS_FCST, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_MI_FCST, header_std, writer_MI_FCST, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Sales_MKT1_FCST, header_std, writer_Sales_MKT1_FCST, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Sales_MKT2_FCST, header_std, writer_Sales_MKT2_FCST, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Sales_MKT3_FCST, header_std, writer_Sales_MKT3_FCST, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Sales_MKT4_FCST, header_std, writer_Sales_MKT4_FCST, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Sales_MKT5_FCST, header_std, writer_Sales_MKT5_FCST, [], [15, 16, 19, 32], [17, 18])

    write_to_excel(data_MKT_ACT, header_std, writer_MKT_ACT, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_MKT_Launch_ACT, header_std, writer_MKT_Launch_ACT, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_APAC_CC_ACT, header_std, writer_APAC_CC_ACT, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Customer_Service_ACT, header_std, writer_Customer_Service_ACT, [], [15, 16, 19, 32], [17, 18])
    show_time(start)
    write_to_excel(data_DTS_ACT, header_std, writer_DTS_ACT, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_MI_ACT, header_std, writer_MI_ACT, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Sales_MKT1_ACT, header_std, writer_Sales_MKT1_ACT, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Sales_MKT2_ACT, header_std, writer_Sales_MKT2_ACT, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Sales_MKT3_ACT, header_std, writer_Sales_MKT3_ACT, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Sales_MKT4_ACT, header_std, writer_Sales_MKT4_ACT, [], [15, 16, 19, 32], [17, 18])
    write_to_excel(data_Sales_MKT5_ACT, header_std, writer_Sales_MKT5_ACT, [], [15, 16, 19, 32], [17, 18])
    os.system('cls')
    print("The program has ended. The program runs for a total of %.2f seconds" % (time.time() - start))
    os.system("pause")

except Exception as e:
    os.system("cls")
    print("The program throws an exception. Please check whether your operation is correct.")
    print("The following is the exception:")
    print("************************************")
    print(e)
    print("************************************")
    os.system("pause")
