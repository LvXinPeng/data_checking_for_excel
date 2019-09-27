# for idx in range(len(agent_code_unique)):
# print(str(agent_code_unique[idx]) + "-三级电气技师-" + str(
#     df[(df['三级电气技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()))
# print(str(agent_code_unique[idx]) + "-三级机械技师-" + str(
#     df[(df['三级机械技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()))
# print(str(agent_code_unique[idx]) + "-三级空调技师-" + str(
#     df[(df['三级空调技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()))
# print(str(agent_code_unique[idx]) + "-四级技师-" + str(
#     df[(df['四级技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()))
# print("==============")
# g3elec = df[(df['三级电气技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
# g3mech = df[(df['三级机械技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
# g3airc = df[(df['三级空调技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
# g4tech = df[(df['四级技师'] == '100%') & (agent_code_unique[idx] == agent_code.values)].index.tolist()
# print(g3elec)
# print(g3mech)
# print(g3airc)
# print(g4tech)
# print("***********************")

#
# print('V700 – V900'.split(' – V')[0].split('V')[1])
# print(len( 'D:/ASP - Erin/Raw Data/ASP Dbase Report 2019 FCST 1+11.xlsx'.split('.')[0][-9:]))
import os

print('            -'.replace('-', '').replace(' ','') == '')
listdir =  os.listdir('D:/ASP - Erin/Raw Data/')
file = 'D:/ASP - Erin/Raw Data/' + listdir[0]
print(file)
print(listdir)
