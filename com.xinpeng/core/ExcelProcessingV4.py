import pandas as pd

# df = pd.read_excel("D:/record-part.xlsx")
df = pd.read_excel("D:/record-all.xlsx")
df_std = pd.read_excel("D:/mapping-old.xlsx")
df_std_new = pd.read_excel("D:/mapping-new.xlsx")
writer = pd.ExcelWriter(path="D:/record-new-test-all.xlsx", mode='w', engine='openpyxl')
col_origin = ['经销商代码', '一级技师', '二级技师', '三级电气技师', '三级机械技师', '三级空调技师', '四级技师', '一级专属服务顾问', '二级专属服务顾问', '保修专员', '高级保修专员',
              '客户服务经理', '混合动力技师', '经销商内训师', '一级钣金技师', '二级钣金技师', '三级钣金技师', '服务管家', '一级配件专员', '二级配件专员', '客户关系经理', '车间管理',
              '一级专属服务技师']
col = ['经销商代码', '能力等级', '一级技师', '二级技师', '三级电气技师', '三级机械技师', '三级空调技师', '四级技师', '一级专属服务顾问', '二级专属服务顾问', '保修专员', '高级保修专员',
       '客户服务经理', '混合动力技师', '经销商内训师', '一级钣金技师	', '二级钣金技师	', '三级钣金技师	', '服务管家', '一级配件专员', '二级配件专员', '客户关系经理',
       '车间管理']
col_new_origin = ['经销商代码', '一级技师', '二级技师', '三级电气技师', '三级机械技师', '三级空调技师', '四级技师', '一级专属服务顾问', '二级专属服务顾问', '保修专员',
                  '高级保修专员', '客户服务经理', '经销商内训师', '一级配件专员', '二级配件专员', '客户服务经理', '一级专属服务技师']
col_new = ['经销商代码', '能力等级', '一级技师', '二级技师', '三级电气技师', '三级机械技师', '三级空调技师', '四级技师', '一级专属服务顾问', '二级专属服务顾问', '保修专员',
           '高级保修专员', '客户服务经理', '经销商内训师', '一级配件专员', '二级配件专员', '客户服务经理']

agent_code = df['经销商代码']
agent_code_std = df_std['经销商代码']
agent_code_std_new = df_std_new['经销商代码']
agent_code_unique = agent_code.unique()

g3_elec = []
g3_mech = []
g3_airc = []
g4_tech = []
rd = []
g3_elec_act = 0
g3_mech_act = 0
g3_airc_act = 0
g4_tech_act = 0
rd_new = []

for idx in range(len(agent_code_unique)):
    g3_elec = df[(df['三级电气技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
    g3_mech = df[(df['三级机械技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
    g3_airc = df[(df['三级空调技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
    g4_tech = df[(df['四级技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
    # 行号索引去重； 行号索引可以表示一个技师
    for a in g4_tech[:]:
        for b in g3_elec[:]:
            temp = b
            if b == a:
                g3_elec.remove(b)
            for bc in g3_mech[:]:
                if bc == temp:
                    g3_mech.remove(bc)
            for bd in g3_airc[:]:
                if bd == temp:
                    g3_airc.remove(bd)
        for c in g3_mech[:]:
            temp = c
            if c == a:
                g3_mech.remove(c)
            for cd in g3_airc[:]:
                if cd == temp:
                    g3_airc.remove(cd)
        for d in g3_airc[:]:
            if d == a:
                g3_airc.remove(d)
    g3_elec_tmp = len(g3_elec)
    g3_mech_tmp = len(g3_mech)
    g3_airc_tmp = len(g3_airc)
    g4_tech_tmp = len(g4_tech)

    # ======技师标准数量==========
    row_data_std = []  # 临时存放一行一年以上经销商的 技师数量标准
    row_data_std_new = []  # 临时存放一行新经销商的 技师数量标准
    # 一年以上经销商
    if agent_code_unique[idx] in agent_code_std.values:
        print(df_std[agent_code_unique[idx] == agent_code_std.values].values[0])
        # 一年以上经销商 技师数量标准
        row_data_std.append(agent_code_unique[idx])
        row_data_std.append('标准')
        for row_idx in range(21):
            row_data_std.append(df_std[agent_code_unique[idx] == agent_code_std.values].values[0][row_idx + 2])
        # 一年以上经销商 三级电气/机械/空调/四级技师 数量标准
        g3_elec_std = df_std[agent_code_unique[idx] == agent_code_std.values].values[0][4]
        g3_mech_std = df_std[agent_code_unique[idx] == agent_code_std.values].values[0][5]
        g3_airc_std = df_std[agent_code_unique[idx] == agent_code_std.values].values[0][6]
        g4_tech_std = df_std[agent_code_unique[idx] == agent_code_std.values].values[0][7]
        # 根据标准 计算三级电气/机械/空调/四级技师的 实际数量
        if g4_tech_tmp <= g4_tech_std:
            g3_elec_act = g3_elec_tmp
            g3_mech_act = g3_mech_tmp
            g3_airc_act = g3_airc_tmp
        else:
            gap4tech = g4_tech_tmp - g4_tech_std
            gap4elec = abs(g3_elec_tmp - g3_elec_std)
            gap4mech = abs(g3_mech_tmp - g3_mech_std)
            gap4airc = abs(g3_airc_tmp - g3_airc_std)
            if g3_elec_tmp >= g3_elec_std and g3_mech_tmp >= g3_mech_std and g3_airc_tmp >= g3_airc_std:
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp
                g3_airc_act = g3_airc_tmp
            elif g3_elec_tmp < g3_elec_std and gap4elec >= gap4tech:
                g3_elec_act = g3_elec_tmp + gap4tech
                g3_mech_act = g3_mech_tmp
                g3_airc_act = g3_airc_tmp
            elif g3_elec_tmp < g3_elec_std and gap4elec < gap4tech:
                if g3_mech_tmp >= g3_mech_std and g3_airc_tmp >= g3_airc_tmp:
                    g3_elec_act = g3_elec_tmp + gap4elec
                    g3_mech_act = g3_mech_tmp
                    g3_airc_act = g3_airc_tmp
                elif g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech >= (gap4tech - gap4elec):
                    g3_elec_act = g3_elec_tmp + gap4elec
                    g3_mech_act = g3_mech_tmp + gap4tech - gap4elec
                    g3_airc_act = g3_airc_tmp
                elif g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech < (
                        gap4tech - gap4elec) and gap4airc >= (gap4tech - gap4elec - gap4mech):
                    g3_elec_act = g3_elec_tmp + gap4elec
                    g3_mech_act = g3_mech_tmp + gap4mech
                    g3_airc_act = g3_airc_tmp + gap4tech - gap4elec - gap4mech
                elif g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech < (
                        gap4tech - gap4elec) and gap4airc < (gap4tech - gap4elec - gap4mech):
                    g3_elec_act = g3_elec_tmp + gap4elec
                    g3_mech_act = g3_mech_tmp + gap4mech
                    g3_airc_act = g3_airc_tmp + gap4airc
            elif g3_elec_tmp >= g3_elec_std and g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech >= gap4tech:
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp + gap4tech
                g3_airc_act = g3_airc_tmp
            elif g3_elec_tmp >= g3_elec_std and g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech < gap4tech and gap4airc >= (
                    gap4tech - gap4mech):
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp + gap4mech
                g3_airc_act = g3_airc_tmp + gap4tech - gap4mech
            elif g3_elec_tmp >= g3_elec_std and g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech < gap4tech and gap4airc < (
                    gap4tech - gap4mech):
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp + gap4mech
                g3_airc_act = g3_airc_tmp + gap4airc
            elif g3_elec_tmp >= g3_elec_std and g3_mech_tmp >= g3_mech_std and g3_airc_tmp < g3_airc_std and gap4airc >= gap4tech:
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp
                g3_airc_act = g3_airc_tmp + gap4tech
            elif g3_elec_tmp >= g3_elec_std and g3_mech_tmp >= g3_mech_std and g3_airc_tmp < g3_airc_std and gap4airc < gap4tech:
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp
                g3_airc_act = g3_airc_tmp + gap4airc
        # g4_tech_act = g4_tech_tmp
    # 新经销商
    elif agent_code_unique[idx] in agent_code_std_new.values:
        print(df_std_new[agent_code_unique[idx] == agent_code_std_new.values].values[0])
        # 新经销商 技师数量标准
        row_data_std_new.append(agent_code_unique[idx])
        row_data_std_new.append('标准')
        for row_idx in range(15):
            row_data_std_new.append(
                df_std_new[agent_code_unique[idx] == agent_code_std_new.values].values[0][row_idx + 2])
        # 新经销商 三级电气/机械/空调/四级技师 数量标准
        g3_elec_std = df_std_new[agent_code_unique[idx] == agent_code_std_new.values].values[0][4]
        g3_mech_std = df_std_new[agent_code_unique[idx] == agent_code_std_new.values].values[0][5]
        g3_airc_std = df_std_new[agent_code_unique[idx] == agent_code_std_new.values].values[0][6]
        g4_tech_std = df_std_new[agent_code_unique[idx] == agent_code_std_new.values].values[0][7]
        # 根据标准 计算三级电气/机械/空调/四级技师的 实际数量
        if g4_tech_tmp <= g4_tech_std:
            g3_elec_act = g3_elec_tmp
            g3_mech_act = g3_mech_tmp
            g3_airc_act = g3_airc_tmp
        else:
            gap4tech = g4_tech_tmp - g4_tech_std
            gap4elec = abs(g3_elec_tmp - g3_elec_std)
            gap4mech = abs(g3_mech_tmp - g3_mech_std)
            gap4airc = abs(g3_airc_tmp - g3_airc_std)
            if g3_elec_tmp >= g3_elec_std and g3_mech_tmp >= g3_mech_std and g3_airc_tmp >= g3_airc_std:
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp
                g3_airc_act = g3_airc_tmp
            elif g3_elec_tmp < g3_elec_std and gap4elec >= gap4tech:
                g3_elec_act = g3_elec_tmp + gap4tech
                g3_mech_act = g3_mech_tmp
                g3_airc_act = g3_airc_tmp
            elif g3_elec_tmp < g3_elec_std and gap4elec < gap4tech:
                if g3_mech_tmp >= g3_mech_std and g3_airc_tmp >= g3_airc_tmp:
                    g3_elec_act = g3_elec_tmp + gap4elec
                    g3_mech_act = g3_mech_tmp
                    g3_airc_act = g3_airc_tmp
                elif g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech >= (gap4tech - gap4elec):
                    g3_elec_act = g3_elec_tmp + gap4elec
                    g3_mech_act = g3_mech_tmp + gap4tech - gap4elec
                    g3_airc_act = g3_airc_tmp
                elif g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech < (
                        gap4tech - gap4elec) and gap4airc >= (gap4tech - gap4elec - gap4mech):
                    g3_elec_act = g3_elec_tmp + gap4elec
                    g3_mech_act = g3_mech_tmp + gap4mech
                    g3_airc_act = g3_airc_tmp + gap4tech - gap4elec - gap4mech
                elif g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech < (
                        gap4tech - gap4elec) and gap4airc < (gap4tech - gap4elec - gap4mech):
                    g3_elec_act = g3_elec_tmp + gap4elec
                    g3_mech_act = g3_mech_tmp + gap4mech
                    g3_airc_act = g3_airc_tmp + gap4airc
            elif g3_elec_tmp >= g3_elec_std and g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech >= gap4tech:
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp + gap4tech
                g3_airc_act = g3_airc_tmp
            elif g3_elec_tmp >= g3_elec_std and g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech < gap4tech and gap4airc >= (
                    gap4tech - gap4mech):
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp + gap4mech
                g3_airc_act = g3_airc_tmp + gap4tech - gap4mech
            elif g3_elec_tmp >= g3_elec_std and g3_mech_tmp < g3_mech_std and g3_airc_tmp < g3_airc_std and gap4mech < gap4tech and gap4airc < (
                    gap4tech - gap4mech):
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp + gap4mech
                g3_airc_act = g3_airc_tmp + gap4airc
            elif g3_elec_tmp >= g3_elec_std and g3_mech_tmp >= g3_mech_std and g3_airc_tmp < g3_airc_std and gap4airc >= gap4tech:
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp
                g3_airc_act = g3_airc_tmp + gap4tech
            elif g3_elec_tmp >= g3_elec_std and g3_mech_tmp >= g3_mech_std and g3_airc_tmp < g3_airc_std and gap4airc < gap4tech:
                g3_elec_act = g3_elec_tmp
                g3_mech_act = g3_mech_tmp
                g3_airc_act = g3_airc_tmp + gap4airc
    g4_tech_act = g4_tech_tmp

    # 计算处理之后的 实际三级电气/机械/空调/四级技师 数量 添加到同一个list
    special_data = []
    special_data.append(g3_elec_act)
    special_data.append(g3_mech_act)
    special_data.append(g3_airc_act)
    special_data.append(g4_tech_act)

    # ======技师实际数量==========
    row_data = []
    row_data.append(agent_code_unique[idx])
    row_data.append('实际')
    # 一年以上经销商
    if agent_code_unique[idx] in agent_code_std.values:
        for index in range(2):
            # 判断值 ('ASA', '100%') 输出值为True or False
            isExist = (agent_code_unique[idx], '100%') in df.groupby([agent_code, df[col_origin[index + 1]]]).groups
            if isExist:
                # 输出整列的匹配数；
                size_col = df.groupby([agent_code, df[col_origin[index + 1]]]).size()[agent_code_unique[idx], '100%']
                row_data.append(size_col)
            else:
                # 如果在判断 不存在('ASA', '100%')这样类似的组合 即认定为不存在符合要求的技师
                row_data.append(0)
        for index in range(3, 7):
            row_data.append(special_data[abs(3 - index)])
        for index in range(7, 22):
            # 判断值 ('ASA', '100%') 输出值为True or False
            isExist = (agent_code_unique[idx], '100%') in df.groupby([agent_code, df[col_origin[index]]]).groups
            if isExist:
                size_col = df.groupby([agent_code, df[col_origin[index]]]).size()[agent_code_unique[idx], '100%']
                row_data.append(size_col)
            else:
                # 如果在判断 不存在('ASA', '100%')这样类似的组合 即认定为不存在符合要求的技师
                row_data.append(0)
        if (agent_code_unique[idx], '100%') in df.groupby([agent_code, df[col_origin[22]]]).groups:
            size_col = df.groupby([agent_code, df[col_origin[22]]]).size()[agent_code_unique[idx], '100%']
            row_data[7] = row_data[7] + size_col
        rd.append(row_data_std)
        rd.append(row_data)
    # 新经销商
    elif agent_code_unique[idx] in agent_code_std_new.values:
        for index in range(2):
            # 判断值 ('ASA', '100%') 输出值为True or False
            isExist = (agent_code_unique[idx], '100%') in df.groupby([agent_code, df[col_new_origin[index + 1]]]).groups
            if isExist:
                # 输出整列的匹配数；
                size_col = df.groupby([agent_code, df[col_new_origin[index + 1]]]).size()[
                    agent_code_unique[idx], '100%']
                row_data.append(size_col)
            else:
                # 如果在判断 不存在('ASA', '100%')这样类似的组合 即认定为不存在符合要求的技师
                row_data.append(0)
        for index in range(3, 7):
            row_data.append(special_data[abs(3 - index)])
        for index in range(7, 16):
            # 判断值 ('ASA', '100%') 输出值为True or False
            isExist = (agent_code_unique[idx], '100%') in df.groupby([agent_code, df[col_new_origin[index]]]).groups
            if isExist:
                size_col = df.groupby([agent_code, df[col_new_origin[index]]]).size()[agent_code_unique[idx], '100%']
                row_data.append(size_col)
            else:
                # 如果在判断 不存在('ASA', '100%')这样类似的组合 即认定为不存在符合要求的技师
                row_data.append(0)
        if (agent_code_unique[idx], '100%') in df.groupby([agent_code, df[col_new_origin[16]]]).groups:
            size_col = df.groupby([agent_code, df[col_new_origin[16]]]).size()[agent_code_unique[idx], '100%']
            row_data[7] = row_data[7] + size_col
        rd_new.append(row_data_std_new)
        rd_new.append(row_data)

frame = pd.DataFrame(rd, columns=col)
frame.to_excel(excel_writer=writer, sheet_name='开业一年以上经销商', index=False, header=True)
frame_new = pd.DataFrame(rd_new, columns=col_new)
frame_new.to_excel(excel_writer=writer, sheet_name='新经销商', index=False, header=True)
writer.save()
writer.close()
print(rd)
print(rd_new)
print('Finished')
