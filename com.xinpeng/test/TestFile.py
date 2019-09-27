import time

import pandas as pd

df_req = pd.read_excel("C:\\Users\\xinplv\\Downloads\\转发__SMM_DB_data_validation-ASP\\DCC 2019.5.1.xlsx")
# ['线索ID', '大区', '小区', '省份', '城市', '经销商代码', '经销商简称', '线索导入账号', '客户登记时间',
#        '线索下发日期', '来源渠道', '市场活动代码', '市场活动', '客户姓名', '客户电话号码', '称谓', '邮箱',
#        '线索意向车型', '线索备注', '最新线索状态', '线索跟进时间', '是否增量', '销售顾问姓名', '销售顾问手机',
#        '潜客ID', '潜客创建日期', '潜客意向等级', '潜客意向车型', '潜客最新状态', '潜客备注', '第一次跟进时间',
#        '第一次跟进方式', '最近一次跟进时间', '最近一次跟进方式', '最后一次跟进备注', '首次到店时间', '二次到店时间',
#        '最近一次到店时间', '潜客跟进次数', '到店次数', '试驾时间', '试驾次数', '报价时间', '战败时间', '战败原因',
#        '订单时间', '开票日期', 'RD日期', '线索首次下发时间', '垂媒销售', '是否转介绍', '介绍人姓名', '转介绍时间',
#        '是否增购', '原沃尔沃车架号', '增购时间', '是否置换', '置换车辆车架号', '置换时间', '备注1', '备注2',
#        '备注3']
df_view = pd.read_excel("C:\\Users\\xinplv\\Desktop\\some results\\dcc-0501.xlsx")
# ['id', 'grpcode', 'groups_name', 'orgcode', 'companys_name',
#        'province_code', 'province', 'city_code', 'city', 'loccode',
#        'code_name', 'loc_name', 'loading_user_id', 'loading_user', 'shop_time',
#        'allot_time', 'source_dd_id', 'source_name', 'activity_code',
#        'activity_name', 'customers_name', 'customers_contact', 'appellation',
#        'email', 'intention_id', 'intention_name', 'clue_remark',
#        'status_dd_id', 'status_name', 'followinfo_time', 'isNew',
#        'accountuser_id', 'ownerrole_name', 'phonenumber', 'customers_id',
#        'cus_createtime', 'cus_level_name', 'cust_intention_name',
#        'cust_status', 'cust_remark', 'clue_allot_number', 'testdrive_time',
#        'testdrive_number', 'quotation_time', 'zhanbai_time', 'zhanbai_remark',
#        'orderinfo_time', 'create_date', 'rdtime', 'createtime', 'mediasale',
#        'introduction_dd_id', 'introduction_name', 'buymore_dd_id',
#        'buymore_oldvin', 'replacement_dd_id', 'replacement_vin',
#        'introduction_date', 'buymore_date', 'replacement_date',
#        'eVolvoleadsID', 'pk_dealer_business', 'pk_dealer_report']

# print(df_req['线索ID'][0])
id_same_count = 0
last_time = time.time()
print(last_time)
for idx_r in df_req['线索ID'][:]:
    for idx_v in df_view['id'][:]:
        if idx_r == idx_v:
            id_same_count = id_same_count + 1
            print(id_same_count)

print("time: " + str(time.time() - last_time))
print("view: " + str(len(df_view['id'])))
print("request: " + str(len(df_req['线索ID'])))
print("same: " + str(id_same_count))

# view: 7395
# request: 7314
# same: 7395
