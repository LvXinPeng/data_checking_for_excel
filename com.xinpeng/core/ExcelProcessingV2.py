import pandas as pd

# df = pd.read_excel("D:/record-part.xlsx")
df = pd.read_excel("D:/record-all.xlsx")
df_std = pd.read_excel("D:/mapping.xlsx")
writer = pd.ExcelWriter(path="D:/record-new.xlsx", mode='w', engine='openpyxl')

agent_code = df['经销商代码']
agent_code_std = df_std['经销商代码']
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

for idx in range(len(agent_code_unique)):
    g3_elec = df[(df['三级电气技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
    g3_mech = df[(df['三级机械技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
    g3_airc = df[(df['三级空调技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
    g4_tech = df[(df['四级技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()

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

    if agent_code_unique[idx] in agent_code_std.values:
        print(df_std[agent_code_unique[idx] == agent_code_std.values].values[0])
        g3_elec_std = df_std[agent_code_unique[idx] == agent_code_std.values].values[0][4]
        g3_mech_std = df_std[agent_code_unique[idx] == agent_code_std.values].values[0][5]
        g3_airc_std = df_std[agent_code_unique[idx] == agent_code_std.values].values[0][6]
        g4_tech_std = df_std[agent_code_unique[idx] == agent_code_std.values].values[0][7]
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

    special_data = []
    special_data.append(g3_elec_act)
    special_data.append(g3_mech_act)
    special_data.append(g3_airc_act)
    special_data.append(g4_tech_act)

    row_data = []
    row_data.append(agent_code_unique[idx])

    for index in range(8):
        # 判断值 ('ASA', '100%') 输出值为True or False
        isExist = (agent_code_unique[idx], '100%') in df.groupby([agent_code, df.columns[index + 5]]).groups
        if isExist:
            # 输出整列的匹配数；
            size_col = df.groupby([agent_code, df.columns[index + 5]]).size()[agent_code_unique[idx], '100%']
            row_data.append(size_col)
        else:
            # 如果在判断 不存在('ASA', '100%')这样类似的组合 即认定为不存在符合要求的技师
            row_data.append(0)

    for index in range(13, 17):
        row_data.append(special_data[abs(13 - index)])

    for index in range(17, 29):
        # 判断值 ('ASA', '100%') 输出值为True or False
        isExist = (agent_code_unique[idx], '100%') in df.groupby([agent_code, df.columns[index]]).groups
        if isExist:
            size_col = df.groupby([agent_code, df.columns[index]]).size()[agent_code_unique[idx], '100%']
            row_data.append(size_col)
        else:
            # 如果在判断 不存在('ASA', '100%')这样类似的组合 即认定为不存在符合要求的技师
            row_data.append(0)
    rd.append(row_data)
col = ['经销商代码', '一级钣金技师	', '一级技师', '沃尔沃入职证书 Volvo Onboard Certification', '二级技师', '二级钣金技师', '三级公共技师', '三级钣金技师',
       '一级专属服务顾问', '三级电气技师', '三级机械技师', '三级空调技师', '四级技师', '二级专属服务顾问', '一级专属服务技师', '客户服务经理', '车间管理', '保修专员', '高级保修专员',
       '一级配件专员', '二级配件专员', '经销商内训师', '服务管家', '客户关系经理', '混合动力技师']
frame = pd.DataFrame(rd, columns=col)
frame.to_excel(writer, index=False, header=True)
writer.save()
writer.close()
print(rd)
print('Finished')
