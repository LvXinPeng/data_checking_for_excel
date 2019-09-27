import datetime
import pandas as pd

writer_old = pd.ExcelWriter(path="D:/mapping-old.xlsx", mode='w', engine='openpyxl')
writer_new = pd.ExcelWriter(path="D:/mapping-new.xlsx", mode='w', engine='openpyxl')
# ['Dlr Code', 'Area', 'Area Manager', '无钣喷业务', 'Format','Operation Date (From DN)']
df_dealer = pd.read_excel("D:/Dealer-Identification-Report.xlsx")
# ['经销商代码','区域',  '区域经理',  201804,  201805,  201806,  201807,  201808, 201809,
# 201810,  201811,  201812,  201901,  201902,  201903]
df_num_avg = pd.read_excel("D:/avg-by-mon.xlsx")
# ['月均入厂台次', '一级技师', '二级技师', '三级电气技师', '三级机械技师', '三级空调技师', '四级技师',
#        '混合动力技师', '一级钣金技师', '二级钣金技师', '三级钣金技师', '一级专属服务顾问', '二级专属服务顾问', '服务管家',
#        '客户关系经理', '客户服务经理', '经销商内训师', '一级配件专员', '二级配件专员', '保修专员', '高级保修专员',
#        '车间管理']
df_stdo = pd.read_excel("D:/std-old.xlsx")
# ['建店规模', '一级技师', '二级技师', '三级电气技师', '三级机械技师', '三级空调技师', '四级技师',
#        '一级专属服务顾问', '二级专属服务顾问', '客户关系经理', '客户服务经理', '经销商内训师', '一级配件专员',
#        '二级配件专员', '保修专员', '高级保修专员']
df_stdn = pd.read_excel("D:/std-new.xlsx")
input_year = input("Please input date like the format '2019-08-20':\n")
data_old = []
data_new = []


def get_months(str1, str2):
    year1 = datetime.datetime.strptime(str1[0:10], "%Y-%m-%d").year
    year2 = datetime.datetime.strptime(str2[0:10], "%Y-%m-%d").year
    month1 = datetime.datetime.strptime(str1[0:10], "%Y-%m-%d").month
    month2 = datetime.datetime.strptime(str2[0:10], "%Y-%m-%d").month
    months = abs((year1 - year2) * 12 + (month1 - month2))
    return months


# 第一层：按dealer表遍历code
for d_idx in range(len(df_dealer['Dlr Code'])):
    row_data_old = []
    row_data_new = []
    # 区分代理商的年限是一年以上还是一下
    if get_months(input_year, str(df_dealer.ix[d_idx, 5])) >= 12:
        # 第二层：按月均表遍历code
        for a_idx in range(len(df_num_avg['经销商代码'])):
            # 判断内层的code是否与外层相同
            if df_dealer['Dlr Code'][d_idx] == df_num_avg['经销商代码'][a_idx]:
                # 内外层code相同，则添加code到row_data数组里
                row_data_old.append(df_dealer['Dlr Code'][d_idx])
                num_sum = 0
                for idx in range(len(df_num_avg.columns) - 3):
                    num_sum = num_sum + int(df_num_avg.ix[a_idx, df_num_avg.columns[idx + 3]])
                num_avg = num_sum / 12
                if 0 < num_avg <= 100:
                    row_data_old.append('L1')
                    for o_idx in range(len(df_stdo['月均入厂台次'])):
                        if df_stdo['月均入厂台次'][o_idx] == '0 - 100':
                            for oc_idx in range(len(df_stdo.columns) - 1):
                                row_data_old.append(df_stdo.ix[o_idx, df_stdo.columns[oc_idx + 1]])
                elif 101 <= num_avg <= 200:
                    row_data_old.append('L2')
                    for o_idx in range(len(df_stdo['月均入厂台次'])):
                        if df_stdo['月均入厂台次'][o_idx] == '101 - 200':
                            for oc_idx in range(len(df_stdo.columns) - 1):
                                row_data_old.append(df_stdo.ix[o_idx, df_stdo.columns[oc_idx + 1]])
                elif 201 <= num_avg <= 300:
                    row_data_old.append('L3')
                    for o_idx in range(len(df_stdo['月均入厂台次'])):
                        if df_stdo['月均入厂台次'][o_idx] == '201 - 300':
                            for oc_idx in range(len(df_stdo.columns) - 1):
                                row_data_old.append(df_stdo.ix[o_idx, df_stdo.columns[oc_idx + 1]])
                elif 301 <= num_avg <= 400:
                    row_data_old.append('L4')
                    for o_idx in range(len(df_stdo['月均入厂台次'])):
                        if df_stdo['月均入厂台次'][o_idx] == '301 - 400':
                            for oc_idx in range(len(df_stdo.columns) - 1):
                                row_data_old.append(df_stdo.ix[o_idx, df_stdo.columns[oc_idx + 1]])
                elif 401 <= num_avg <= 500:
                    row_data_old.append('L5')
                    for o_idx in range(len(df_stdo['月均入厂台次'])):
                        if df_stdo['月均入厂台次'][o_idx] == '401 - 500':
                            for oc_idx in range(len(df_stdo.columns) - 1):
                                row_data_old.append(df_stdo.ix[o_idx, df_stdo.columns[oc_idx + 1]])
                elif 501 <= num_avg <= 600:
                    row_data_old.append('L6')
                    for o_idx in range(len(df_stdo['月均入厂台次'])):
                        if df_stdo['月均入厂台次'][o_idx] == '501 - 600':
                            for oc_idx in range(len(df_stdo.columns) - 1):
                                row_data_old.append(df_stdo.ix[o_idx, df_stdo.columns[oc_idx + 1]])
                elif 601 <= num_avg <= 700:
                    row_data_old.append('L7')
                    for o_idx in range(len(df_stdo['月均入厂台次'])):
                        if df_stdo['月均入厂台次'][o_idx] == '601 - 700':
                            for oc_idx in range(len(df_stdo.columns) - 1):
                                row_data_old.append(df_stdo.ix[o_idx, df_stdo.columns[oc_idx + 1]])
                elif 701 <= num_avg <= 800:
                    row_data_old.append('L8')
                    for o_idx in range(len(df_stdo['月均入厂台次'])):
                        if df_stdo['月均入厂台次'][o_idx] == '701 - 800':
                            for oc_idx in range(len(df_stdo.columns) - 1):
                                row_data_old.append(df_stdo.ix[o_idx, df_stdo.columns[oc_idx + 1]])
                elif 801 <= num_avg <= 900:
                    row_data_old.append('L9')
                    for o_idx in range(len(df_stdo['月均入厂台次'])):
                        if df_stdo['月均入厂台次'][o_idx] == '801 - 900':
                            for oc_idx in range(len(df_stdo.columns) - 1):
                                row_data_old.append(df_stdo.ix[o_idx, df_stdo.columns[oc_idx + 1]])
                elif 901 <= num_avg <= 1000:
                    row_data_old.append('L10')
                    for o_idx in range(len(df_stdo['月均入厂台次'])):
                        if df_stdo['月均入厂台次'][o_idx] == '901 - 1000':
                            for oc_idx in range(len(df_stdo.columns) - 1):
                                row_data_old.append(df_stdo.ix[o_idx, df_stdo.columns[oc_idx + 1]])
                elif 1001 <= num_avg <= 1800:
                    row_data_old.append('L11')
                    for o_idx in range(len(df_stdo['月均入厂台次'])):
                        if df_stdo['月均入厂台次'][o_idx] == '1001 - 1800':
                            for oc_idx in range(len(df_stdo.columns) - 1):
                                row_data_old.append(df_stdo.ix[o_idx, df_stdo.columns[oc_idx + 1]])
        data_old.append(row_data_old)
    else:
        row_data_new.append(df_dealer['Dlr Code'][d_idx])
        # 第二层：按标准表遍历建店规模
        for n_idx in range(len(df_stdn['建店规模'])):
            if df_dealer['Format'][d_idx] == df_stdn['建店规模'][n_idx]:
                row_data_new.append(df_dealer['Format'][d_idx])
                for nc_idx in range(len(df_stdn.columns) - 1):
                    row_data_new.append(df_stdn.ix[n_idx, df_stdn.columns[nc_idx + 1]])
            elif df_dealer['Format'][d_idx] == '<V150' and df_stdn['建店规模'][n_idx] == 'V100':
                row_data_new.append(df_dealer['Format'][d_idx])
                for nc_idx in range(len(df_stdn.columns) - 1):
                    row_data_new.append(df_stdn.ix[0, df_stdn.columns[nc_idx + 1]])
            elif df_dealer['Format'][d_idx] == 'V700' and df_stdn['建店规模'][n_idx] == 'V700 – V900':
                row_data_new.append(df_dealer['Format'][d_idx])
                for nc_idx in range(len(df_stdn.columns) - 1):
                    row_data_new.append(df_stdn.ix[4, df_stdn.columns[nc_idx + 1]])
            elif df_dealer['Format'][d_idx] == 'V900' and df_stdn['建店规模'][n_idx] == 'V700 – V900':
                row_data_new.append(df_dealer['Format'][d_idx])
                for nc_idx in range(len(df_stdn.columns) - 1):
                    row_data_new.append(df_stdn.ix[4, df_stdn.columns[nc_idx + 1]])
            elif df_dealer['Format'][d_idx] == 'V1200' and df_stdn['建店规模'][n_idx] == 'V1200 – V1800':
                row_data_new.append(df_dealer['Format'][d_idx])
                for nc_idx in range(len(df_stdn.columns) - 1):
                    row_data_new.append(df_stdn.ix[5, df_stdn.columns[nc_idx + 1]])
            elif df_dealer['Format'][d_idx] == 'V1800' and df_stdn['建店规模'][n_idx] == 'V1200 – V1800':
                row_data_new.append(df_dealer['Format'][d_idx])
                for nc_idx in range(len(df_stdn.columns) - 1):
                    row_data_new.append(df_stdn.ix[5, df_stdn.columns[nc_idx + 1]])
        data_new.append(row_data_new)
print(data_old)
print(data_new)
col_old = ['经销商代码', 'LEVEL', '一级技师', '二级技师', '三级电气技师', '三级机械技师', '三级空调技师', '四级技师','混合动力技师', '一级钣金技师', '二级钣金技师', '三级钣金技师', '一级专属服务顾问', '二级专属服务顾问', '服务管家','客户关系经理', '客户服务经理', '经销商内训师', '一级配件专员', '二级配件专员', '保修专员', '高级保修专员','车间管理']
frame_old = pd.DataFrame(data_old, columns=col_old)
frame_old.to_excel(writer_old, index=False, header=True)
writer_old.save()
writer_old.close()
col_new = ['经销商代码', 'LEVEL', '一级技师', '二级技师', '三级电气技师', '三级机械技师', '三级空调技师', '四级技师',
           '一级专属服务顾问', '二级专属服务顾问', '客户关系经理', '客户服务经理', '经销商内训师', '一级配件专员',
           '二级配件专员', '保修专员', '高级保修专员']
frame_new = pd.DataFrame(data_new, columns=col_new)
frame_new.to_excel(writer_new, index=False, header=True)
writer_new.save()
writer_new.close()
print('Finished')
