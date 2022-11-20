# coding=utf-8
# 窗口提示
print('请确认“数据源”文件夹内：',
      '最新销售台账是否准备完毕。')
print('本次运行时间大约1分钟,请勿关闭本窗口！')
print('已完成{:.0%}'.format(0))
#-*- coding : utf-8-*-
# coding:unicode_escape
# 记录运行时间
#导入使用库
import datetime
import re
import time
import pandas as pd
import xlsxwriter
import warnings
import numpy as np
import os
import shutil
from tkinter import filedialog
current_file_path = os.getcwd()
os.chdir(current_file_path)
warnings.filterwarnings('ignore')
start_time = time.time()
#加载源数据
filePath1 = '数据源'
file_name1 = os.listdir(filePath1)
for i in range(len(file_name1)):
    if str(file_name1[i]).count('~$') == 0:
        report_all = pd.read_excel(filePath1 + '/' + str(file_name1[i]),sheet_name='一.总表明细',header=1)
        report_out = pd.read_excel(filePath1 + '/' + str(file_name1[i]), sheet_name='2.出货')[['订单序号',"实际出货日期","实际数量","出货序号"]]
        report_order = pd.read_excel(filePath1 + '/' + str(file_name1[i]), sheet_name='1、订单')[['订单序号', "质保期"]]
        report_receive =pd.read_excel(filePath1 + '/' + str(file_name1[i]), sheet_name='3.验收')[['订单序号',"系统录入日期",'终验收时间','是否预验收','数量',"验收序号"]]
        report_card = pd.read_excel(filePath1 + '/' + str(file_name1[i]), sheet_name='5.发票')[['发票序号', "开票日期"]]
report_all =report_all.drop(0).reset_index(drop=True)
report_all =report_all[report_all['回款完成率']<1].reset_index(drop=True)
report_out =report_out.drop(0).reset_index(drop=True)
report_receive =report_receive.drop(0).reset_index(drop=True)
del report_all['实际数量']
del report_all['数量']
#del report_all['质保期']
report_out = report_out.rename(columns={'出货序号':'子序号'})
report_card  = report_card.rename(columns={'发票序号':'子序号'})

#处理订单表字段，增加质保期限字段
report_order['质保期']=report_order['质保期'].fillna('空值')
report_order=report_order[report_order['质保期']!='空值'].reset_index(drop=True)
report_order['质保期']=report_order['质保期'].astype(str)
report_order['质保期限']=0
report_order.loc[(report_order['质保期'].str.contains("一年|1年|365|12个月")),'质保期限'] = 12
report_order.loc[(report_order['质保期'].str.contains("3个月|90|0.25年")),'质保期限'] = 3
report_order.loc[(report_order['质保期'].str.contains("6个月|180|半年|0.5年")),'质保期限'] =6
report_order.loc[(report_order['质保期'].str.contains("730|2年|24个月|24")),'质保期限'] = 24
report_order.loc[(report_order['质保期'].str.contains("18个月")),'质保期限'] = 18
report_order.loc[(report_order['质保期'].str.contains("1年质保期")),'质保期限'] = 12
report_order.loc[(report_order['质保期'].str.contains("3年|36个月")),'质保期限'] = 36
report_order.loc[(report_order['质保期'].str.contains("一年|1年|365|12个月")),'质保期限'] = 12
report_order.loc[(report_order['质保期'].str.contains("一个月|1个月")),'质保期限'] = 1
#处理验收表，将系统录入时间和终验收时间合并成一个字段
default_date = '1990/01/01'
report_receive['系统录入日期']=report_receive['系统录入日期'].fillna(default_date)
report_receive['终验收时间']=report_receive['终验收时间'].fillna(default_date)
report_receive.loc[(report_receive['终验收时间']==default_date)&(report_receive['是否预验收']!="Y"),'终验收时间'] = report_receive['系统录入日期']
del report_receive['系统录入日期']

#合并出货表的实际出货日期
report_all_1=pd.merge(report_all,report_out,left_on='序号',right_on='订单序号',how='left')
#合并验收表的终验收时间	是否预验收
report_all_2=pd.merge(report_all_1,report_receive,left_on='子序号',right_on='验收序号',how='left')
#合并订单表的质保期
report=pd.merge(report_all_2,report_order,left_on='序号',right_on='订单序号',how='left')
#合并发票开票日期
report=pd.merge(report,report_card,on='子序号',how='left')
#排下序
report['序号']=report['序号'].astype('string')
report=report.sort_values(by=['序号','实际出货日期']).reset_index(drop=True)
report = report.rename(columns={'出货完成率':'实际出货完成率'})
report = report.rename(columns={'验收完成率':'实际验收完成率'})

#选择需要字段
need=['序号','子序号', '接单据点',  '订单签订日期',
       '公司名称', '公司简称', '主订单号',  '订单名称', '订单单位', '订单数量',
       '税率/汇率', '币别',  '含税金额', 'USD',  '结款方式',
       '结款方式说明', '预付款付款天数', '出货款付款天数', '验收款付款天数', '质保款付款天数', '业务员',
       '项目号',  '立项名称', '立项数量',  '大项目名称', '产品线',
       '项目类型',  '实际数量', '实际未税金额', 'USD.1', '实际出货完成率', '数量',
       '未税金额.1', '含税金额.1', 'USD.2', '实际验收完成率', '回款原币金额', '回款完成率', '开票总金额',
       '开票完成率','开票日期', '实际出货日期', '终验收时间', '是否预验收', '质保期', '质保期限']
#处理report格式
for col in report.columns:
       if col not in need:
              del report[col]

#锁定表字段
report = report.reindex(columns=['序号','子序号', '接单据点',  '订单签订日期',
       '公司名称', '公司简称', '主订单号',  '订单名称', '订单单位', '订单数量',
       '税率/汇率', '币别',  '含税金额', 'USD',  '结款方式',
       '结款方式说明', '预付款付款天数', '出货款付款天数', '验收款付款天数', '质保款付款天数','质保期', '质保期限','质保时间', '业务员',
       '项目号',  '立项名称', '立项数量',  '大项目名称', '产品线',
       '项目类型',  '实际数量', '实际未税金额', 'USD.1', '实际出货完成率', '实际出货日期','数量',
       '未税金额.1', '含税金额.1', 'USD.2', '实际验收完成率','终验收时间', '是否预验收', '回款原币金额', '回款完成率', '开票总金额',
       '开票完成率','开票日期',
       '当前状态', '预付款占比','预付款应付', '预付款欠款','预付款应付时间','预付款账龄','预付款当前欠款',
        '发货款占比', '发货款应付','发货款欠款','发货款应付时间','发货款账龄','收货款当前欠款',
        '验收款占比',  '验收款应付','验收款欠款','验收款应付时间','验收款账龄','验收款当前欠款',
        '质保款占比', '质保款应付',  '质保款欠款','质保款应付时间','质保款账龄','质保款当前欠款','当前欠款总计'])
#日期填充
report['质保时间']=default_date
date_col=['订单签订日期','实际出货日期', '终验收时间','预付款应付时间','发货款应付时间','验收款应付时间','质保款应付时间',"质保时间",'开票日期']
report[date_col]=report[date_col].fillna(default_date)
#文本
str_col=['质保期','大项目名称','是否预验收','税率/汇率',"子序号",'结款方式说明','立项名称','订单名称','项目类型','订单单位','业务员','项目号']
report[str_col] = report[str_col].fillna('').astype(str)


#数字
int_col=[ '订单数量',   '含税金额', 'USD','结款方式',
        '预付款付款天数', '出货款付款天数', '验收款付款天数', '质保款付款天数',
        '立项数量',     '实际数量', '实际未税金额', 'USD.1', '实际出货完成率', '数量',
       '未税金额.1', '含税金额.1', 'USD.2', '实际验收完成率', '回款原币金额', '回款完成率', '开票总金额',
       '开票完成率',  '是否预验收', '质保期', '质保期限','预付款占比','预付款应付', '预付款欠款', '发货款占比', '发货款应付'
        ,'发货款欠款', '验收款占比',  '验收款应付','验收款欠款',
        '质保款占比', '质保款应付',  '质保款欠款']
report[int_col]=report[int_col].fillna(0)
#日期设定
report['终验收时间'] = pd.to_datetime(report["终验收时间"],errors='coerce')
report['实际出货日期'] = pd.to_datetime(report["实际出货日期"],errors='coerce')
report['订单签订日期'] = pd.to_datetime(report['订单签订日期'],errors='coerce')
report[date_col]=report[date_col].fillna(default_date)
report['订单签订日期'] = pd.to_datetime(report['订单签订日期'],errors='coerce')
report['开票日期'] = pd.to_datetime(report['开票日期'],errors='coerce')
#质保时间写入数据并规范格式
for i in range(len(report)):
    if report['终验收时间'][i] != pd.Timestamp(1990, 1, 1) and  report['质保期限'][i] !=0:
        report['质保时间'][i]=report['终验收时间'][i]+pd.DateOffset(months=report['质保期限'][i])
    else:
        next
report['质保时间'] = pd.to_datetime(report["质保时间"],errors='coerce')

#插入自定义函数给当前状态使用
def current_statu(out_date, pecent_rece, rece_date, yes_rece,prote_date):     #['出货日期']['验收完成率'],['验收时间],['是否预验收'],['质保时间'],['已入库量'])
    temp = []
    length = len(out_date)
    for i in range(length):
        if out_date[i]==pd.Timestamp(1990, 1, 1):
            temp.append('未出货')
        else:
            if pecent_rece[i]<1:
                temp.append('已出货未验收')
            else:
                if rece_date[i]==pd.Timestamp(1990, 1, 1) and yes_rece[i]=='Y'  :
                    temp.append('预验收')
                else:
                    if datetime.datetime.today().date()<prote_date[i]:
                        temp.append('质保期内')
                    else:
                        temp.append('已过质保期')
    return temp

report['当前状态']=''
report.loc[:, '当前状态'] = current_statu(report['实际出货日期'], report['实际验收完成率'],report['终验收时间'],report['是否预验收']
                                            ,report['质保时间'] )
report.loc[(report['质保期限']==0) & (report['当前状态']=='已过质保期'),'当前状态'] = '已验收无质保期'
#计算实际出货完成率
report.loc[(report['立项数量']!=0),'实际出货完成率'] =report['实际数量']/ report['立项数量']
report.loc[(report['立项数量']!=0),'实际验收完成率'] =report['数量']/ report['立项数量']

#计算各款项占比

for i in range(len(report)):
    if len(str(report['结款方式'][i]))==2:
        c=int(str(report['结款方式'][i])[0])
        d=int(str(report['结款方式'][i])[1])
        report['验收款占比'][i]=c*0.1
        report['质保款占比'][i]=d*0.1
    if len(str(report['结款方式'][i]))==3:
        b=int(str(report['结款方式'][i])[0])
        c=int(str(report['结款方式'][i])[1])
        d=int(str(report['结款方式'][i])[2])
        report['发货款占比'][i]=b*0.1
        report['验收款占比'][i]=c*0.1
        report['质保款占比'][i]=d*0.1
    if len(str(report['结款方式'][i]))==4:
        a=int(str(report['结款方式'][i])[0])
        b=int(str(report['结款方式'][i])[1])
        c=int(str(report['结款方式'][i])[2])
        d=int(str(report['结款方式'][i])[3])
        report['预付款占比'][i]=a*0.1
        report['发货款占比'][i]=b*0.1
        report['验收款占比'][i]=c*0.1
        report['质保款占比'][i]=d*0.1
    if len(str(report['结款方式'][i]))>4:
        next
report.loc[(report['结款方式']==10),'验收款占比'] = 1
report.loc[(report['结款方式']==100),'发货款占比'] = 1
report.loc[(report['结款方式']==1000),'预付款占比'] = 1

report.loc[(report['结款方式']=='(2.5)(3.5)(3)(1)'),'预付款占比'] = 0.25
report.loc[(report['结款方式']=='(2.5)(3.5)(3)(1)'),'发货款占比'] = 0.35
report.loc[(report['结款方式']=='(2.5)(3.5)(3)(1)'),'验收款占比'] = 0.3
report.loc[(report['结款方式']=='(2.5)(3.5)(3)(1)'),'质保款占比'] = 0.1

report.loc[(report['结款方式']=='53(1.5)(0.5)'),'预付款占比'] = 0.5
report.loc[(report['结款方式']=='53(1.5)(0.5)'),'发货款占比'] = 0.3
report.loc[(report['结款方式']=='53(1.5)(0.5)'),'验收款占比'] = 0.15
report.loc[(report['结款方式']=='53(1.5)(0.5)'),'质保款占比'] = 0.05

report.loc[(report['结款方式']=='323(1.5)(0.5)'),'预付款占比'] = 0.3
report.loc[(report['结款方式']=='323(1.5)(0.5)'),'发货款占比'] = 0.2
report.loc[(report['结款方式']=='323(1.5)(0.5)'),'验收款占比'] = 0.45
report.loc[(report['结款方式']=='323(1.5)(0.5)'),'质保款占比'] = 0.05

report.loc[(report['结款方式']=='603(0.5)(0.5)'),'预付款占比'] = 0.6
report.loc[(report['结款方式']=='603(0.5)(0.5)'),'发货款占比'] = 0
report.loc[(report['结款方式']=='603(0.5)(0.5)'),'验收款占比'] = 0.3
report.loc[(report['结款方式']=='603(0.5)(0.5)'),'质保款占比'] = 0.1

report.loc[(report['结款方式']=='243(0.5)(0.3)(0.2)'),'预付款占比'] = 0.2
report.loc[(report['结款方式']=='243(0.5)(0.3)(0.2)'),'发货款占比'] = 0.4
report.loc[(report['结款方式']=='243(0.5)(0.3)(0.2)'),'验收款占比'] = 0.3
report.loc[(report['结款方式']=='243(0.5)(0.3)(0.2)'),'质保款占比'] = 0.1

report.loc[(report['结款方式']=='306(0.5)(0.5)'),'预付款占比'] = 0.3
report.loc[(report['结款方式']=='306(0.5)(0.5)'),'发货款占比'] = 0
report.loc[(report['结款方式']=='306(0.5)(0.5)'),'验收款占比'] = 0.6
report.loc[(report['结款方式']=='306(0.5)(0.5)'),'质保款占比'] = 0.1

report.loc[(report['结款方式']=='333(0.5)(0.5)'),'预付款占比'] = 0.3
report.loc[(report['结款方式']=='333(0.5)(0.5)'),'发货款占比'] = 0
report.loc[(report['结款方式']=='333(0.5)(0.5)'),'验收款占比'] = 0.6
report.loc[(report['结款方式']=='333(0.5)(0.5)'),'质保款占比'] = 0.1
#(3.805)(2.655)(2.655)(0.885) (2.5)(3.5)31 (3.1)(1.5)1(3.4)1 (1.15)(2.655)(2.655)(2.655）（0.885） (5.2)(2.1)(2.1)(0.7)
report.loc[(report['结款方式']=='(3.805)(2.655)(2.655)(0.885)'),'预付款占比'] = 0.3805
report.loc[(report['结款方式']=='(3.805)(2.655)(2.655)(0.885)'),'发货款占比'] = 0.2655
report.loc[(report['结款方式']=='(3.805)(2.655)(2.655)(0.885)'),'验收款占比'] = 0.2655
report.loc[(report['结款方式']=='(3.805)(2.655)(2.655)(0.885)'),'质保款占比'] = 0.885

report.loc[(report['结款方式']=='(3.1)(1.5)1(3.4)1'),'预付款占比'] = 0.31
report.loc[(report['结款方式']=='(3.1)(1.5)1(3.4)1'),'发货款占比'] = 0.15
report.loc[(report['结款方式']=='(3.1)(1.5)1(3.4)1'),'验收款占比'] = 0.34
report.loc[(report['结款方式']=='(3.1)(1.5)1(3.4)1'),'质保款占比'] = 0.1

report.loc[(report['结款方式']=='(2.5)(3.5)31'),'预付款占比'] = 0.25
report.loc[(report['结款方式']=='(2.5)(3.5)31'),'发货款占比'] = 0.35
report.loc[(report['结款方式']=='(2.5)(3.5)31'),'验收款占比'] = 0.3
report.loc[(report['结款方式']=='(2.5)(3.5)31'),'质保款占比'] = 0.1
report['结款方式']=report['结款方式'].fillna("人生")

report.loc[(report['结款方式']=='(5.2)(2.1)(2.1)(0.7)'),'预付款占比'] = 0.52
report.loc[(report['结款方式']=='(5.2)(2.1)(2.1)(0.7)'),'发货款占比'] = 0.21
report.loc[(report['结款方式']=='(5.2)(2.1)(2.1)(0.7)'),'验收款占比'] = 0.21
report.loc[(report['结款方式']=='(5.2)(2.1)(2.1)(0.7)'),'质保款占比'] = 0.07

report.loc[(report['结款方式']=='(1.15)(2.655)(2.655)(2.655）（0.885）'),'预付款占比'] = 0.115
report.loc[(report['结款方式']=='(1.15)(2.655)(2.655)(2.655）（0.885）'),'发货款占比'] = 0.2655
report.loc[(report['结款方式']=='(1.15)(2.655)(2.655)(2.655）（0.885）'),'验收款占比'] = 0.531
report.loc[(report['结款方式']=='(1.15)(2.655)(2.655)(2.655）（0.885）'),'质保款占比'] = 0.0885

for i in range(len(report)):
    if str(report.loc[i,'结款方式']).count('-')>1:
        report.loc[i, '预付款占比']=float(report.loc[i, '结款方式'].split('-')[0])*0.1  #float(a.split('-')[1])
        report.loc[i, '发货款占比'] = float(report.loc[i, '结款方式'].split('-')[1])*0.1
        report.loc[i, '验收款占比'] = float(report.loc[i, '结款方式'].split('-')[2])*0.1
        report.loc[i, '质保款占比'] = float(report.loc[i, '结款方式'].split('-')[3])*0.1

#计算各款项应付

report.loc[(report['币别']=='CNY'),'预付款应付'] = report['含税金额']*report['预付款占比']
report.loc[(report['币别']=='CNY')&(report['立项数量']!=0),'发货款应付'] = report['含税金额']*report['发货款占比']*report['实际数量']/report['立项数量']
#report.loc[(report['立项数量']==0),'发货款应付'] = 0
report.loc[(report['币别']=='CNY')&(report['立项数量']!=0),'验收款应付'] = report['含税金额']*report['验收款占比']*report['数量']/report['立项数量']
report.loc[(report['币别']=='CNY')&(report['立项数量']!=0),'质保款应付'] = report['含税金额']*report['质保款占比']*report['数量']/report['立项数量']


report.loc[(report['币别']=='USD'),'预付款应付'] = report['USD']*report['预付款占比']
report.loc[(report['币别']=='USD'),'发货款应付'] = (report['USD']*report['发货款占比']*report['实际数量']/report['立项数量'])
report.loc[(report['币别']=='USD'),'验收款应付'] = (report['USD']*report['验收款占比']*report['数量']/report['立项数量'])
report.loc[(report['币别']=='USD'),'质保款应付'] = (report['USD']*report['质保款占比']*report['数量']/report['立项数量'])
report.loc[(report['立项数量']==0),'发货款应付'] = 0
report.loc[(report['立项数量']==0),'验收款应付'] = 0
report.loc[(report['立项数量']==0),'质保款应付'] = 0
#计算各款项欠款
report.loc[(report['回款原币金额']>=report['预付款应付']),'预付款欠款'] = 0
report.loc[(report['回款原币金额']<report['预付款应付']),'预付款欠款'] =report['预付款应付']-report['回款原币金额']

report.loc[(report['回款原币金额']>=(report['发货款应付']+report['预付款应付'])),'发货款欠款'] = 0
report.loc[(report['回款原币金额']<(report['发货款应付']+report['预付款应付'])),'发货款欠款'] =report['发货款应付']+report['预付款应付']-report['回款原币金额']

report.loc[(report['回款原币金额']>=(report['发货款应付']+report['预付款应付']+report['验收款应付'])),'验收款欠款'] = 0
report.loc[(report['回款原币金额']<(report['发货款应付']+report['预付款应付']+report['验收款应付'])),'验收款欠款']=report['发货款应付']+report['预付款应付']+report['验收款应付']-report['回款原币金额']

report.loc[(report['回款原币金额']>=(report['发货款应付']+report['预付款应付']+report['验收款应付']+report['质保款应付'])),'质保款欠款'] = 0
report.loc[(report['回款原币金额']<(report['发货款应付']+report['预付款应付']+report['验收款应付']+report['质保款应付'])),'质保款欠款']=report['质保款应付']+report['发货款应付']+report['预付款应付']+report['验收款应付']-report['回款原币金额']



#计算分批物料的回款分配
#加入辅助列 数量-统计 实际数量-统计 发货款全部应付
group1=report.groupby(['序号']).agg({"数量":"sum"}).add_suffix('-统计').reset_index()
report=pd.merge(report,group1,on='序号',how='left')
group2=report.groupby(['序号']).agg({"实际数量":"sum"}).add_suffix('-统计').reset_index()
report=pd.merge(report,group2,on='序号',how='left')
report.loc[(report['币别']=='CNY'),'发货款全部应付'] = report['含税金额']*report['发货款占比']*report['实际数量-统计']/report['立项数量']
report.loc[(report['币别']=='CNY'),'验收款全部应付'] = report['含税金额']*report['验收款占比']*report['数量-统计']/report['立项数量']
report.loc[(report['币别']=='CNY'),'质保款全部应付'] = report['含税金额']*report['质保款占比']*report['数量-统计']/report['立项数量']
report.loc[(report['币别']=='USD'),'发货款全部应付'] = (report['USD']*report['发货款占比']*report['实际数量-统计']/report['立项数量'])
report.loc[(report['币别']=='USD'),'验收款全部应付'] = (report['USD']*report['验收款占比']*report['数量-统计']/report['立项数量'])
report.loc[(report['币别']=='USD'),'质保款全部应付'] = (report['USD']*report['质保款占比']*report['数量-统计']/report['立项数量'])
#分批
report['发货回款余额'] = report['回款原币金额'] - report['预付款应付']
# report.loc[(report['发货回款余额']<=0),'发货回款余额'] = 0
for i in range(1, len(report)):
    if report["序号"][i] == report["序号"][i - 1]:
        if report['发货回款余额'][i - 1] <= 1:
            report['发货回款余额'][i - 1] = 0
            report['发货回款余额'][i] = 0
            report['发货款欠款'][i - 1] = report['发货款应付'][i-1]
            report['发货款欠款'][i] = report['发货款应付'][i]
        if report['发货回款余额'][i - 1] > 1:
            report['发货款欠款'][i - 1] = report['发货款应付'][i - 1] - report['发货回款余额'][i - 1]
            report['发货回款余额'][i] = report['发货回款余额'][i - 1] - report['发货款应付'][i - 1]
            if report['发货款应付'][i - 1] >= report['发货回款余额'][i - 1]:
                report['发货回款余额'][i] = 0
                report['发货款欠款'][i] = report['发货款应付'][i]
            if report['发货款应付'][i - 1] < report['发货回款余额'][i - 1]:
                report['发货款欠款'][i - 1] = 0
                report['发货款欠款'][i] = report['发货款应付'][i] - report['发货回款余额'][i]
                if report['发货款应付'][i] < report['发货回款余额'][i]:
                    report['发货款欠款'][i] = 0

report['验收回款余额'] = report['回款原币金额'] - report['预付款应付'] - report['发货款全部应付']
# report.loc[(report['验收回款余额']<=0),'验收回款余额'] = 0
for i in range(1, len(report)):
    if report["序号"][i] == report["序号"][i - 1]:
        if report['验收回款余额'][i - 1] <= 1:
            report['验收回款余额'][i - 1] = 0
            report['验收回款余额'][i] = 0
            report['验收款欠款'][i - 1] = report['验收款应付'][i-1]
            report['验收款欠款'][i] = report['验收款应付'][i]
        if report['验收回款余额'][i - 1] > 1:
            report['验收款欠款'][i - 1] = report['验收款应付'][i - 1] - report['验收回款余额'][i - 1]
            report['验收回款余额'][i] = report['验收回款余额'][i - 1] - report['验收款应付'][i - 1]
            if report['验收款应付'][i - 1] >= report['验收回款余额'][i - 1]:
                report['验收回款余额'][i] = 0
                report['验收款欠款'][i] = report['验收款应付'][i]
            if report['验收款应付'][i - 1] < report['验收回款余额'][i - 1]:
                report['验收款欠款'][i - 1] = 0
                report['验收款欠款'][i] = report['验收款应付'][i] - report['验收回款余额'][i]
                if report['验收款应付'][i] < report['验收回款余额'][i]:
                    report['验收款欠款'][i] = 0

report['质保回款余额'] = report['回款原币金额'] - report['预付款应付'] - report['发货款全部应付'] - report['验收款全部应付']
# report.loc[(report['验收回款余额']<=0),'验收回款余额'] = 0
# report.loc[(report['质保回款余额']<=0),'质保回款余额'] = 0
for i in range(1, len(report)):
    if report["序号"][i] == report["序号"][i - 1]:
        if report['质保回款余额'][i - 1] <= 1:
            report['质保回款余额'][i - 1] = 0
            report['质保回款余额'][i] = 0
            report['质保款欠款'][i - 1] = report['质保款应付'][i-1]
            report['质保款欠款'][i] = report['质保款应付'][i]
        if report['质保回款余额'][i - 1] > 1:
            report['质保款欠款'][i - 1] = report['质保款应付'][i - 1] - report['质保回款余额'][i - 1]
            report['质保回款余额'][i] = report['质保回款余额'][i - 1] - report['质保款应付'][i - 1]
            if report['质保款应付'][i - 1] >= report['质保回款余额'][i - 1]:
                report['质保回款余额'][i] = 0
                report['质保款欠款'][i] = report['质保款应付'][i]
            if report['质保款应付'][i - 1] < report['质保回款余额'][i - 1]:
                report['质保款欠款'][i - 1] = 0
                report['质保款欠款'][i] = report['质保款应付'][i] - report['质保回款余额'][i]
                if report['质保款应付'][i] < report['质保回款余额'][i]:
                    report['质保款欠款'][i] = 0

report.loc[(report['发货款欠款']>=report['发货款应付']),'发货款欠款'] =report['发货款应付']
report.loc[(report['验收款欠款']>=report['验收款应付']),'验收款欠款'] =report['验收款应付']
report.loc[(report['质保款欠款']>=report['质保款应付']),'质保款欠款'] =report['质保款应付']
report.loc[(report['预付款欠款']<1)&(report['预付款欠款']>0),'预付款欠款'] =0
report.loc[(report['发货款欠款']<1)&(report['发货款欠款']>0),'发货款欠款'] =0
report.loc[(report['验收款欠款']<1)&(report['验收款欠款']>0),'验收款欠款'] =0
report.loc[(report['质保款欠款']<1)&(report['质保款欠款']>0),'质保款欠款'] =0

report.loc[report['订单签订日期']!=pd.Timestamp(1990, 1, 1),'预付款应付时间']=pd.to_datetime(report['订单签订日期'])+pd.to_timedelta(report['预付款付款天数'],unit='D')
report.loc[report['实际出货日期']!=pd.Timestamp(1990, 1, 1),'发货款应付时间']=pd.to_datetime(report['实际出货日期'])+pd.to_timedelta(report['出货款付款天数'],unit='D')
report.loc[report['终验收时间']!=pd.Timestamp(1990, 1, 1),'验收款应付时间']=pd.to_datetime(report['终验收时间'])+pd.to_timedelta(report['验收款付款天数'],unit='D')
report.loc[report['质保时间']!=pd.Timestamp(1990, 1, 1),'质保款应付时间']=pd.to_datetime(report['质保时间'])+pd.to_timedelta(report['质保款付款天数'],unit='D')

report.loc[(report['立项数量']==0),'质保款欠款'] =0
report.loc[(report['立项数量']==0),'验收款欠款'] =0
report.loc[(report['立项数量']==0),'发货款欠款'] =0
report.loc[(report['立项数量']==0),'质保款应收'] =0
report.loc[(report['立项数量']==0),'验收款应收'] =0
report.loc[(report['立项数量']==0),'发货款应收'] =0

#账龄计算

report['预付款应付时间'] = pd.to_datetime(report["预付款应付时间"],errors='coerce')
report['发货款应付时间'] = pd.to_datetime(report["发货款应付时间"],errors='coerce')
report['验收款应付时间'] = pd.to_datetime(report["验收款应付时间"],errors='coerce')
report['质保款应付时间'] = pd.to_datetime(report["质保款应付时间"],errors='coerce')
time_delta1 = list(datetime.datetime.today() - report['预付款应付时间'])
report.loc[:, '预付款账龄'] = [item.days if item.days<10000 else 0 for item in time_delta1]
time_delta2 = list(datetime.datetime.today() - report['发货款应付时间'])
report.loc[:, '发货款账龄'] = [item.days if item.days<10000 else 0 for item in time_delta2]
time_delta3 = list(datetime.datetime.today() - report['验收款应付时间'])
report.loc[:, '验收款账龄'] = [item.days if item.days<10000 else 0 for item in time_delta3]
time_delta4 = list(datetime.datetime.today() - report['质保款应付时间'])
report.loc[:, '质保款账龄'] = [item.days if item.days<10000 else 0 for item in time_delta4]
#当前欠款
report['预付款当前欠款']=0
report['发货款当前欠款']=0
report['验收款当前欠款']=0
report['质保款当前欠款']=0
report.loc[(report['预付款账龄']>=0),'预付款当前欠款']=report['预付款欠款']

report.loc[(report['发货款账龄']>=0)&(report["发货款应付时间"]!= pd.Timestamp(1990, 1, 1)),'发货款当前欠款']=report['发货款欠款']
#report.loc[(report['实际数量']<=0),'发货款当前欠款']=report['发货款欠款']
report.loc[(report['验收款账龄']>=0)&(report["验收款应付时间"]!= pd.Timestamp(1990, 1, 1)),'验收款当前欠款']=report['验收款欠款']
#report.loc[(report['数量']<=0),'验收款当前欠款']=report['验收款欠款']
report.loc[(report['质保款账龄']>=0)&(report["质保款应付时间"]!= pd.Timestamp(1990, 1, 1)),'质保款当前欠款']=report['质保款欠款']
#report.loc[(report['数量']<=0),'验收款当前欠款']=report['验收款当前欠款']
#当前欠款总计
group2=report.groupby(['序号']).agg({"发货款当前欠款":"sum"}).add_suffix('统计').reset_index()
report=pd.merge(report,group2,on='序号',how='left')
group3=report.groupby(['序号']).agg({"验收款当前欠款":"sum"}).add_suffix('统计').reset_index()
report=pd.merge(report,group3,on='序号',how='left')
group4=report.groupby(['序号']).agg({"质保款当前欠款":"sum"}).add_suffix('统计').reset_index()
report=pd.merge(report,group4,on='序号',how='left')
report['当前欠款总计']=report['预付款当前欠款']+report['发货款当前欠款统计']+report['验收款当前欠款统计']+report['质保款当前欠款统计']
del report['发货款当前欠款统计']
del report['验收款当前欠款统计']
del report['质保款当前欠款统计']
#
report['小项目名称']=report['大项目名称']
report.loc[(report['小项目名称']==''),'小项目名称'] = '其他'
report['预付款当前应付']=0
report['发货款当前应付']=0
report['验收款当前应付']=0
report['质保款当前应付']=0
report.loc[(report['预付款账龄']>0),'预付款当前应付']=report['预付款应付']
report.loc[(report['发货款账龄']>0),'发货款当前应付']=report['发货款应付']
report.loc[(report['验收款账龄']>0),'验收款当前应付']=report['验收款应付']
report.loc[(report['质保款账龄']>0),'质保款当前应付']=report['质保款应付']

group2_1=report.groupby(['序号']).agg({"发货款当前应付":"sum"}).add_suffix('统计').reset_index()
report=pd.merge(report,group2_1,on='序号',how='left')
group3_1=report.groupby(['序号']).agg({"验收款当前应付":"sum"}).add_suffix('统计').reset_index()
report=pd.merge(report,group3_1,on='序号',how='left')
group4_1=report.groupby(['序号']).agg({"质保款当前应付":"sum"}).add_suffix('统计').reset_index()
report=pd.merge(report,group4_1,on='序号',how='left')

report['当前实付结余']=report['回款原币金额']-report['预付款当前应付']-report['发货款当前应付统计']-report['验收款当前应付统计']-report['质保款当前应付统计']
report.loc[(report['当前实付结余']<0.1),'当前实付结余']=0
del report['预付款当前应付']
del report['发货款当前应付']
del report['验收款当前应付']
del report['质保款当前应付']

del report['发货款当前应付统计']
del report['验收款当前应付统计']
del report['质保款当前应付统计']
#预付款单表

reportstart=report[['序号','小项目名称','公司简称','预付款当前欠款','预付款账龄']]
reportstart=reportstart.drop_duplicates(subset='序号',keep='first').reset_index(drop=True)
reportstartall=reportstart.groupby(['公司简称']).agg({"预付款当前欠款":"sum"}).add_suffix('统计').reset_index()
reportstart1=reportstart[(reportstart['预付款账龄']<=90)&(reportstart['预付款账龄']>=0)].groupby(['公司简称']).agg({"预付款当前欠款":"sum"}).add_suffix('统计1').reset_index()
reportstart2=reportstart[(reportstart['预付款账龄']<=180)&(reportstart['预付款账龄']>90)].groupby(['公司简称']).agg({"预付款当前欠款":"sum"}).add_suffix('统计2').reset_index()
reportstart3=reportstart[(reportstart['预付款账龄']<=365)&(reportstart['预付款账龄']>180)].groupby(['公司简称']).agg({"预付款当前欠款":"sum"}).add_suffix('统计3').reset_index()
reportstart4=reportstart[(reportstart['预付款账龄']<=730)&(reportstart['预付款账龄']>365)].groupby(['公司简称']).agg({"预付款当前欠款":"sum"}).add_suffix('统计4').reset_index()
reportstart5=reportstart[(reportstart['预付款账龄']<=1095)&(reportstart['预付款账龄']>730)].groupby(['公司简称']).agg({"预付款当前欠款":"sum"}).add_suffix('统计5').reset_index()
reportstart6=reportstart[(reportstart['预付款账龄']<=10000)&(reportstart['预付款账龄']>1095)].groupby(['公司简称']).agg({"预付款当前欠款":"sum"}).add_suffix('统计6').reset_index()

report_start=pd.merge(reportstartall,reportstart1,on='公司简称',how='left')
report_start=pd.merge(report_start,reportstart2,on='公司简称',how='left')
report_start=pd.merge(report_start,reportstart3,on='公司简称',how='left')
report_start=pd.merge(report_start,reportstart4,on='公司简称',how='left')
report_start=pd.merge(report_start,reportstart5,on='公司简称',how='left')
report_start=pd.merge(report_start,reportstart6,on='公司简称',how='left')
del report_start['预付款当前欠款统计']
report_start=report_start.rename(columns={'预付款当前欠款统计1':'90天内欠款统计','预付款当前欠款统计2':'90-180天内欠款统计','预付款当前欠款统计3':'180-365天内欠款统计','预付款当前欠款统计4':'1-2年欠款统计','预付款当前欠款统计5':'2-3年欠款统计','预付款当前欠款统计6':'大于3年欠款统计'})
strt_int=['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计']
report_start[strt_int]=report_start[strt_int].fillna(0)
report_start['预付款应收']=report_start['90天内欠款统计']+report_start['90-180天内欠款统计']+report_start['180-365天内欠款统计']+report_start['1-2年欠款统计']+report_start['2-3年欠款统计']+report_start['大于3年欠款统计']
report_start=report_start.sort_values(by=['预付款应收'],axis=0,ascending=False).reset_index(drop=True)

reporta=report_start.copy() #做客户分表用
reporta = reporta.rename(columns={'预付款应收':'汇总'})
reporta['类型']='预付款'

reportstart['预付款当前欠款1']=reportstart['预付款当前欠款']/10000
reportstart['预付款当前欠款1']=reportstart['预付款当前欠款1'].round(2)
reportstart['范围1']=''
reportstart.loc[(reportstart['预付款账龄']>=0) &(reportstart['预付款账龄']<=90),'范围1'] = '90天内'
reportstart.loc[(reportstart['预付款账龄']>90) &(reportstart['预付款账龄']<=180),'范围1'] = '90-180天内'
reportstart.loc[(reportstart['预付款账龄']>180) &(reportstart['预付款账龄']<=365),'范围1'] = '180-365天内'
reportstart.loc[(reportstart['预付款账龄']>365) &(reportstart['预付款账龄']<=730),'范围1'] = '1-2年'
reportstart.loc[(reportstart['预付款账龄']>730) &(reportstart['预付款账龄']<=1095),'范围1'] = '2-3年'
reportstart.loc[(reportstart['预付款账龄']>1095) &(reportstart['预付款账龄']<=10000),'范围1'] = '大于3年'
len1=len(report_start)
if len1>10:
    report_start.loc['Col_sum'] =report_start.iloc[10:len1,1:8].sum(axis=0)
    report_start.loc['Col_sum','公司简称']='其他'
    report_start=report_start.drop(report_start.index[10:len1])
    report_start=report_start.reset_index(drop=True)
report_start[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','预付款应收']]=report_start[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','预付款应收']]/10000
report_start[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','预付款应收']]=report_start[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','预付款应收']].round(2)
report_start['账款说明']=''
for i in range(len(report_start)-1):
    report_time1=reportstart[reportstart['公司简称']==report_start['公司简称'][i]].sort_values(by=['预付款当前欠款'],axis=0,ascending=False).reset_index(drop=True)
    #合并求和同范围和项目名称的给到说明
    report_time1 = report_time1.groupby(['小项目名称', '范围1'])["预付款当前欠款1"].sum().add_suffix('').reset_index().sort_values(by=['预付款当前欠款1'],axis=0,ascending=False).reset_index(drop=True)
    if len(report_time1)>=3:
        report_start['账款说明'][i]=str(report_time1['范围1'][0])+': '+str(report_time1['小项目名称'][0])+'---'+str(report_time1['预付款当前欠款1'][0])+'    '+str(report_time1['范围1'][1])+': '+str(report_time1['小项目名称'][1])+'---'+str(report_time1['预付款当前欠款1'][1])+'    '+str(report_time1['范围1'][2])+': '+str(report_time1['小项目名称'][2])+'---'+str(report_time1['预付款当前欠款1'][2])
    if len(report_time1) < 3:
        report_start['账款说明'][i] = str(report_time1['范围1'][0]) + ': ' + str(report_time1['小项目名称'][0]) + '---' + str(
            report_time1['预付款当前欠款1'][0])

#发货款单表
reportooutt=report[['公司简称','小项目名称','发货款当前欠款','发货款账龄']]
#reportooutt=reportooutt.drop_duplicates(subset='序号',keep='first').reset_index(drop=True)
reportoouttall=reportooutt.groupby(['公司简称']).agg({"发货款当前欠款":"sum"}).add_suffix('统计').reset_index()
reportooutt1=reportooutt[(reportooutt['发货款账龄']<=90)&(reportooutt['发货款账龄']>=0)].groupby(['公司简称']).agg({"发货款当前欠款":"sum"}).add_suffix('统计1').reset_index()
reportooutt2=reportooutt[(reportooutt['发货款账龄']<=180)&(reportooutt['发货款账龄']>90)].groupby(['公司简称']).agg({"发货款当前欠款":"sum"}).add_suffix('统计2').reset_index()
reportooutt3=reportooutt[(reportooutt['发货款账龄']<=365)&(reportooutt['发货款账龄']>180)].groupby(['公司简称']).agg({"发货款当前欠款":"sum"}).add_suffix('统计3').reset_index()
reportooutt4=reportooutt[(reportooutt['发货款账龄']<=1095)&(reportooutt['发货款账龄']>365)].groupby(['公司简称']).agg({"发货款当前欠款":"sum"}).add_suffix('统计4').reset_index()
reportooutt5=reportooutt[(reportooutt['发货款账龄']<=1095)&(reportooutt['发货款账龄']>730)].groupby(['公司简称']).agg({"发货款当前欠款":"sum"}).add_suffix('统计5').reset_index()
reportooutt6=reportooutt[(reportooutt['发货款账龄']<=10000)&(reportooutt['发货款账龄']>1095)].groupby(['公司简称']).agg({"发货款当前欠款":"sum"}).add_suffix('统计6').reset_index()

report_outt=pd.merge(reportoouttall,reportooutt1,on='公司简称',how='left')
report_outt=pd.merge(report_outt,reportooutt2,on='公司简称',how='left')
report_outt=pd.merge(report_outt,reportooutt3,on='公司简称',how='left')
report_outt=pd.merge(report_outt,reportooutt4,on='公司简称',how='left')
report_outt=pd.merge(report_outt,reportooutt5,on='公司简称',how='left')
report_outt=pd.merge(report_outt,reportooutt6,on='公司简称',how='left')
del report_outt['发货款当前欠款统计']
report_outt=report_outt.rename(columns={'发货款当前欠款统计1':'90天内欠款统计','发货款当前欠款统计2':'90-180天内欠款统计','发货款当前欠款统计3':'180-365天内欠款统计','发货款当前欠款统计4':'1-2年欠款统计','发货款当前欠款统计5':'2-3年欠款统计','发货款当前欠款统计6':'大于3年欠款统计'})
strt_int=['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计']
report_outt[strt_int]=report_outt[strt_int].fillna(0)
report_outt['发货款应收']=report_outt['90天内欠款统计']+report_outt['90-180天内欠款统计']+report_outt['180-365天内欠款统计']+report_outt['1-2年欠款统计']+report_outt['2-3年欠款统计']+report_outt['大于3年欠款统计']
report_outt=report_outt.sort_values(by=['发货款应收'],axis=0,ascending=False).reset_index(drop=True)
reportb=report_outt.copy()
reportb = reportb.rename(columns={'发货款应收':'汇总'})
reportb['类型']='发货款'


reportooutt['发货款当前欠款1']=reportooutt['发货款当前欠款']/10000
reportooutt['发货款当前欠款1']=reportooutt['发货款当前欠款1'].round(2)
reportooutt['范围2']=''
reportooutt.loc[(reportooutt['发货款账龄']>=0) &(reportooutt['发货款账龄']<=90),'范围2'] = '90天内'
reportooutt.loc[(reportooutt['发货款账龄']>90) &(reportooutt['发货款账龄']<=180),'范围2'] = '90-180天内'
reportooutt.loc[(reportooutt['发货款账龄']>180) &(reportooutt['发货款账龄']<=365),'范围2'] = '180-365天内'
reportooutt.loc[(reportooutt['发货款账龄']>365) &(reportooutt['发货款账龄']<=730),'范围2'] = '1-2年'
reportooutt.loc[(reportooutt['发货款账龄']>730) &(reportooutt['发货款账龄']<=1095),'范围2'] = '2-3年'
reportooutt.loc[(reportooutt['发货款账龄']>1095) &(reportooutt['发货款账龄']<=10000),'范围2'] = '大于3年'
len2=len(report_outt)
if len2>10:
    report_outt.loc['Col_sum'] =report_outt.iloc[10:len2,1:8].sum(axis=0)
    report_outt.loc['Col_sum','公司简称']='其他'
    report_outt=report_outt.drop(report_outt.index[10:len2])
    report_outt=report_outt.reset_index(drop=True)
report_outt[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','发货款应收']]=report_outt[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','发货款应收']]/10000
report_outt[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','发货款应收']]=report_outt[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','发货款应收']].round(2)

report_outt['账款说明']=''
for i in range(len(report_outt)-1):
    report_time2=reportooutt[reportooutt['公司简称']==report_outt['公司简称'][i]].sort_values(by=['发货款当前欠款'],axis=0,ascending=False).reset_index(drop=True)
    # 合并求和同范围和项目名称的给到说明
    report_time2 = report_time2.groupby(['小项目名称', '范围2'])["发货款当前欠款1"].sum().add_suffix('').reset_index().sort_values(by=['发货款当前欠款1'],axis=0,ascending=False).reset_index(drop=True)
    if len(report_time2)>=3:
        report_outt['账款说明'][i]=str(report_time2['范围2'][0])+': '+str(report_time2['小项目名称'][0])+'---'+str(report_time2['发货款当前欠款1'][0])+'    '+str(report_time2['范围2'][1])+': '+str(report_time2['小项目名称'][1])+'---'+str(report_time2['发货款当前欠款1'][1])+'    '+str(report_time2['范围2'][2])+': '+str(report_time2['小项目名称'][2])+'---'+str(report_time2['发货款当前欠款1'][2])
    if len(report_time2) < 3:
        report_outt['账款说明'][i] = str(report_time2['范围2'][0]) + ': ' + str(report_time2['小项目名称'][0]) + '---' + str(
            report_time2['发货款当前欠款1'][0])


#验收款单表
reportrece=report[['公司简称','小项目名称','验收款当前欠款','验收款账龄']]
#report_rece=report_rece.drop_duplicates(subset='序号',keep='first').reset_index(drop=True)
report_receall=reportrece.groupby(['公司简称']).agg({"验收款当前欠款":"sum"}).add_suffix('统计').reset_index()
report_rece1=reportrece[(reportrece['验收款账龄']<=90)&(reportrece['验收款账龄']>=0)].groupby(['公司简称']).agg({"验收款当前欠款":"sum"}).add_suffix('统计1').reset_index()
report_rece2=reportrece[(reportrece['验收款账龄']<=180)&(reportrece['验收款账龄']>90)].groupby(['公司简称']).agg({"验收款当前欠款":"sum"}).add_suffix('统计2').reset_index()
report_rece3=reportrece[(reportrece['验收款账龄']<=365)&(reportrece['验收款账龄']>180)].groupby(['公司简称']).agg({"验收款当前欠款":"sum"}).add_suffix('统计3').reset_index()
report_rece4=reportrece[(reportrece['验收款账龄']<=730)&(reportrece['验收款账龄']>365)].groupby(['公司简称']).agg({"验收款当前欠款":"sum"}).add_suffix('统计4').reset_index()
report_rece5=reportrece[(reportrece['验收款账龄']<=1095)&(reportrece['验收款账龄']>730)].groupby(['公司简称']).agg({"验收款当前欠款":"sum"}).add_suffix('统计5').reset_index()
report_rece6=reportrece[(reportrece['验收款账龄']<=10000)&(reportrece['验收款账龄']>1095)].groupby(['公司简称']).agg({"验收款当前欠款":"sum"}).add_suffix('统计6').reset_index()

report_rece=pd.merge(report_receall,report_rece1,on='公司简称',how='left')
report_rece=pd.merge(report_rece,report_rece2,on='公司简称',how='left')
report_rece=pd.merge(report_rece,report_rece3,on='公司简称',how='left')
report_rece=pd.merge(report_rece,report_rece4,on='公司简称',how='left')
report_rece=pd.merge(report_rece,report_rece5,on='公司简称',how='left')
report_rece=pd.merge(report_rece,report_rece6,on='公司简称',how='left')
del report_rece['验收款当前欠款统计']
report_rece=report_rece.rename(columns={'验收款当前欠款统计1':'90天内欠款统计','验收款当前欠款统计2':'90-180天内欠款统计','验收款当前欠款统计3':'180-365天内欠款统计','验收款当前欠款统计4':'1-2年欠款统计','验收款当前欠款统计5':'2-3年欠款统计','验收款当前欠款统计6':'大于3年欠款统计'})
strt_int=['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计']
report_rece[strt_int]=report_rece[strt_int].fillna(0)
report_rece['验收款应收']=report_rece['90天内欠款统计']+report_rece['90-180天内欠款统计']+report_rece['180-365天内欠款统计']+report_rece['1-2年欠款统计']+report_rece['2-3年欠款统计']+report_rece['大于3年欠款统计']
report_rece=report_rece.sort_values(by=['验收款应收'],axis=0,ascending=False).reset_index(drop=True)
reportc=report_rece.copy()
reportc = reportc.rename(columns={'验收款应收':'汇总'})
reportc['类型']='验收款'


reportrece['验收款当前欠款1']=reportrece['验收款当前欠款']/10000
reportrece['验收款当前欠款1']=reportrece['验收款当前欠款1'].round(2)
reportrece['范围3']=''
reportrece.loc[(reportrece['验收款账龄']>=0) &(reportrece['验收款账龄']<=90),'范围3'] = '90天内'
reportrece.loc[(reportrece['验收款账龄']>90) &(reportrece['验收款账龄']<=180),'范围3'] = '90-180天内'
reportrece.loc[(reportrece['验收款账龄']>180) &(reportrece['验收款账龄']<=365),'范围3'] = '180-365天内'
reportrece.loc[(reportrece['验收款账龄']>365) &(reportrece['验收款账龄']<=730),'范围3'] = '1-2年'
reportrece.loc[(reportrece['验收款账龄']>730) &(reportrece['验收款账龄']<=1095),'范围3'] = '2-3年'
reportrece.loc[(reportrece['验收款账龄']>1095) &(reportrece['验收款账龄']<=10000),'范围3'] = '大于3年'
len3=len(report_rece)
if len3>10:
    report_rece.loc['Col_sum'] =report_rece.iloc[10:len3,1:8].sum(axis=0)
    report_rece.loc['Col_sum','公司简称']='其他'
    report_rece=report_rece.drop(report_rece.index[10:len3])
    report_rece=report_rece.reset_index(drop=True)
report_rece[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','验收款应收']]=report_rece[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','验收款应收']]/10000
report_rece[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','验收款应收']]=report_rece[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','验收款应收']].round(2)

report_rece['账款说明']=''
for i in range(len(report_rece)-1):
    report_time3=reportrece[reportrece['公司简称']==report_rece['公司简称'][i]].sort_values(by=['验收款当前欠款'],axis=0,ascending=False).reset_index(drop=True)
    # 合并求和同范围和项目名称的给到说明
    report_time3= report_time3.groupby(['小项目名称', '范围3'])["验收款当前欠款1"].sum().add_suffix('').reset_index().sort_values(by=['验收款当前欠款1'],axis=0,ascending=False).reset_index(drop=True)
    if len(report_time3)>=3:
        report_rece['账款说明'][i]=str(report_time3['范围3'][0])+': '+str(report_time3['小项目名称'][0])+'---'+str(report_time3['验收款当前欠款1'][0])+'    '+str(report_time3['范围3'][1])+': '+str(report_time3['小项目名称'][1])+'---'+str(report_time3['验收款当前欠款1'][1])+'    '+str(report_time3['范围3'][2])+': '+str(report_time3['小项目名称'][2])+'---'+str(report_time3['验收款当前欠款1'][2])
    if len(report_time3) < 3:
        report_rece['账款说明'][i] = str(report_time3['范围3'][0]) + ': ' + str(report_time3['小项目名称'][0]) + '---' + str(
            report_time3['验收款当前欠款1'][0])

#质保款单表
reportpro=report[['公司简称','小项目名称','质保款当前欠款','质保款账龄']]
#report_pro=report_pro.drop_duplicates(subset='序号',keep='first').reset_index(drop=True)
report_proall=reportpro.groupby(['公司简称']).agg({"质保款当前欠款":"sum"}).add_suffix('统计').reset_index()
report_pro1=reportpro[(reportpro['质保款账龄']<=90)&(reportpro['质保款账龄']>=0)].groupby(['公司简称']).agg({"质保款当前欠款":"sum"}).add_suffix('统计1').reset_index()
report_pro2=reportpro[(reportpro['质保款账龄']<=180)&(reportpro['质保款账龄']>90)].groupby(['公司简称']).agg({"质保款当前欠款":"sum"}).add_suffix('统计2').reset_index()
report_pro3=reportpro[(reportpro['质保款账龄']<=365)&(reportpro['质保款账龄']>180)].groupby(['公司简称']).agg({"质保款当前欠款":"sum"}).add_suffix('统计3').reset_index()
report_pro4=reportpro[(reportpro['质保款账龄']<=730)&(reportpro['质保款账龄']>365)].groupby(['公司简称']).agg({"质保款当前欠款":"sum"}).add_suffix('统计4').reset_index()
report_pro5=reportpro[(reportpro['质保款账龄']<=1095)&(reportpro['质保款账龄']>730)].groupby(['公司简称']).agg({"质保款当前欠款":"sum"}).add_suffix('统计5').reset_index()
report_pro6=reportpro[(reportpro['质保款账龄']<=10000)&(reportpro['质保款账龄']>1095)].groupby(['公司简称']).agg({"质保款当前欠款":"sum"}).add_suffix('统计6').reset_index()

report_pro=pd.merge(report_proall,report_pro1,on='公司简称',how='left')
report_pro=pd.merge(report_pro,report_pro2,on='公司简称',how='left')
report_pro=pd.merge(report_pro,report_pro3,on='公司简称',how='left')
report_pro=pd.merge(report_pro,report_pro4,on='公司简称',how='left')
report_pro=pd.merge(report_pro,report_pro5,on='公司简称',how='left')
report_pro=pd.merge(report_pro,report_pro6,on='公司简称',how='left')
del report_pro['质保款当前欠款统计']
report_pro=report_pro.rename(columns={'质保款当前欠款统计1':'90天内欠款统计','质保款当前欠款统计2':'90-180天内欠款统计','质保款当前欠款统计3':'180-365天内欠款统计','质保款当前欠款统计4':'1-2年欠款统计','质保款当前欠款统计5':'2-3年欠款统计','质保款当前欠款统计6':'大于3年欠款统计'})
strt_int=['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计']
report_pro[strt_int]=report_pro[strt_int].fillna(0)
report_pro['质保款应收']=report_pro['90天内欠款统计']+report_pro['90-180天内欠款统计']+report_pro['180-365天内欠款统计']+report_pro['1-2年欠款统计']+report_pro['2-3年欠款统计']+report_pro['大于3年欠款统计']
report_pro=report_pro.sort_values(by=['质保款应收'],axis=0,ascending=False).reset_index(drop=True)
reportd=report_pro.copy()
reportd = reportd.rename(columns={'质保款应收':'汇总'})
reportd['类型']='质保款'

reportpro['质保款当前欠款1']=reportpro['质保款当前欠款']/10000
reportpro['质保款当前欠款1']=reportpro['质保款当前欠款1'].round(2)
reportpro['范围4']=''
reportpro.loc[(reportpro['质保款账龄']>=0) &(reportpro['质保款账龄']<=90),'范围4'] = '90天内'
reportpro.loc[(reportpro['质保款账龄']>90) &(reportpro['质保款账龄']<=180),'范围4'] = '90-180天内'
reportpro.loc[(reportpro['质保款账龄']>180) &(reportpro['质保款账龄']<=365),'范围4'] = '180-365天内'
reportpro.loc[(reportpro['质保款账龄']>365) &(reportpro['质保款账龄']<=730),'范围4'] = '1-2年'
reportpro.loc[(reportpro['质保款账龄']>730) &(reportpro['质保款账龄']<=1095),'范围4'] = '2-3年'
reportpro.loc[(reportpro['质保款账龄']>1095) &(reportpro['质保款账龄']<=10000),'范围4'] = '大于3年'

len4=len(report_pro)
if len4>10:
    report_pro.loc['Col_sum'] =report_pro.iloc[10:len4,1:8].sum(axis=0)
    report_pro.loc['Col_sum','公司简称']='其他'
    report_pro=report_pro.drop(report_pro.index[10:len4])
    report_pro=report_pro.reset_index(drop=True)
report_pro[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','质保款应收']]=report_pro[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','质保款应收']]/10000
report_pro[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','质保款应收']]=report_pro[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','质保款应收']].round(2)

report_pro['账款说明']=''
for i in range(len(report_pro)-1):
    report_time4=reportpro[reportpro['公司简称']==report_pro['公司简称'][i]].sort_values(by=['质保款当前欠款'],axis=0,ascending=False).reset_index(drop=True)
    # 合并求和同范围和项目名称的给到说明
    report_time4 = report_time4.groupby(['小项目名称', '范围4'])["质保款当前欠款1"].sum().add_suffix('').reset_index().sort_values(by=['质保款当前欠款1'],axis=0,ascending=False).reset_index(drop=True)
    if len(report_time4)>=3:
        report_pro['账款说明'][i]=str(report_time4['范围4'][0])+': '+str(report_time4['小项目名称'][0])+'---'+str(report_time4['质保款当前欠款1'][0])+'    '+str(report_time4['范围4'][1])+': '+str(report_time4['小项目名称'][1])+'---'+str(report_time4['质保款当前欠款1'][1])+'    '+str(report_time4['范围4'][2])+': '+str(report_time4['小项目名称'][2])+'---'+str(report_time4['质保款当前欠款1'][2])
    if len(report_time4) < 3:
        report_pro['账款说明'][i] = str(report_time4['范围4'][0]) + ': ' + str(report_time4['小项目名称'][0]) + '---' + str(
            report_time4['质保款当前欠款1'][0])

#总统计表
#预付
report_start_1=report_start.copy()
report_start_1= report_start_1.rename(columns={'公司简称': '类型','预付款应收': '统计'})
del report_start_1['账款说明']
len1=len(report_start_1)
report_start_1.loc['汇总'] =report_start_1.iloc[0:len1,1:8].sum(axis=0)
report_start_1.loc['汇总','类型']='预付款'
report_start_1=report_start_1.drop(report_start_1.index[0:len1])
#发货
report_outt_1=report_outt.copy()
report_outt_1= report_outt_1.rename(columns={'公司简称': '类型','发货款应收': '统计'})
del report_outt_1['账款说明']
len1=len(report_outt_1)
report_outt_1.loc['汇总'] =report_outt_1.iloc[0:len1,1:8].sum(axis=0)
report_outt_1.loc['汇总','类型']='发货款'
report_outt_1=report_outt_1.drop(report_outt_1.index[0:len1])

#验收
report_rece_1=report_rece.copy()
report_rece_1= report_rece_1.rename(columns={'公司简称': '类型','验收款应收': '统计'})
del report_rece_1['账款说明']
len1=len(report_rece_1)
report_rece_1.loc['汇总'] =report_rece_1.iloc[0:len1,1:8].sum(axis=0)
report_rece_1.loc['汇总','类型']='验收款'
report_rece_1=report_rece_1.drop(report_rece_1.index[0:len1])
#质保
report_pro_1=report_pro.copy()
report_pro_1= report_pro_1.rename(columns={'公司简称': '类型','质保款应收': '统计'})
del report_pro_1['账款说明']
len1=len(report_pro_1)
report_pro_1.loc['汇总'] =report_pro_1.iloc[0:len1,1:8].sum(axis=0)
report_pro_1.loc['汇总','类型']='质保款'
report_pro_1=report_pro_1.drop(report_pro_1.index[0:len1])

#合并
report_to = pd.concat([report_start_1, report_outt_1]).reset_index(drop=True) # 合并2张表格
report_to=pd.concat([report_to, report_rece_1]).reset_index(drop=True) # 合并2张表格
report_to=pd.concat([report_to, report_pro_1]).reset_index(drop=True) # 合并2张表格
report_to.loc['汇总'] =report_to.iloc[0:len1,1:8].sum(axis=0)
report_to.loc['汇总','类型']='统计'


# 总客户表
report_cus = pd.concat([reporta, reportb]).reset_index(drop=True) # 合并2张表格
report_cus = pd.concat([report_cus, reportc]).reset_index(drop=True) # 合并2张表格
report_cus = pd.concat([report_cus, reportd]).reset_index(drop=True) # 合并2张表格
report_cus = report_cus.reindex(columns=['公司简称','类型','90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','汇总'])
report_cus[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','汇总']]=report_cus[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','汇总']]/10000
report_cus[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','汇总']]=report_cus[['90天内欠款统计','90-180天内欠款统计','180-365天内欠款统计','1-2年欠款统计','2-3年欠款统计','大于3年欠款统计','汇总']].round(2)
report_cus = report_cus.rename(columns={'汇总':'单项汇总'})
cus=report_cus.groupby(['公司简称']).agg({"单项汇总":"sum"}).add_suffix('-统计').reset_index()
report_cus=pd.merge(report_cus,cus,on='公司简称',how='left')
report_cus = report_cus.rename(columns={'单项汇总-统计':'汇总'})
report_cus=report_cus.sort_values(by=['汇总','公司简称'],axis=0,ascending=False).reset_index(drop=True)
#设定格式
report = report.reindex(columns=['序号','子序号', '接单据点',  '订单签订日期',
       '公司名称', '公司简称', '主订单号',  '订单名称', '订单单位', '订单数量',
       '税率/汇率', '币别',  '含税金额', 'USD',  '结款方式',
       '结款方式说明', '预付款付款天数', '出货款付款天数', '验收款付款天数', '质保款付款天数','质保期', '质保期限','质保时间', '业务员',
       '项目号',  '立项名称', '立项数量',  '大项目名称', '产品线',
       '项目类型',  '实际数量', '实际未税金额', 'USD.1', '实际出货完成率', '实际出货日期','数量',
       '未税金额.1', '含税金额.1', 'USD.2', '实际验收完成率','终验收时间', '是否预验收', '回款原币金额', '回款完成率', '开票总金额',
       '开票完成率','开票日期',
        '当前状态', '预付款占比','预付款应付', '预付款欠款','预付款应付时间','预付款账龄','预付款当前欠款',
        '发货款占比', '发货款应付','发货款欠款','发货款应付时间','发货款账龄','发货款当前欠款',
        '验收款占比',  '验收款应付','验收款欠款','验收款应付时间','验收款账龄','验收款当前欠款',
        '质保款占比', '质保款应付',  '质保款欠款','质保款应付时间','质保款账龄','质保款当前欠款','当前欠款总计','当前实付结余'])
#清理格式方便写入excel '预付款应付时间','发货款应付时间','验收款应付时间','质保款应付时间'.dt.strftime('%Y-%m-%d')
report['终验收时间'] = pd.to_datetime(report["终验收时间"],errors='coerce').dt.strftime('%Y-%m-%d')
report['实际出货日期'] = pd.to_datetime(report["实际出货日期"],errors='coerce').dt.strftime('%Y-%m-%d')
report['订单签订日期'] = pd.to_datetime(report['订单签订日期'],errors='coerce').dt.strftime('%Y-%m-%d')
report['预付款应付时间'] = pd.to_datetime(report["预付款应付时间"],errors='coerce').dt.strftime('%Y-%m-%d')
report['发货款应付时间'] = pd.to_datetime(report["发货款应付时间"],errors='coerce').dt.strftime('%Y-%m-%d')
report['验收款应付时间'] = pd.to_datetime(report["验收款应付时间"],errors='coerce').dt.strftime('%Y-%m-%d')
report['质保款应付时间'] = pd.to_datetime(report["质保款应付时间"],errors='coerce').dt.strftime('%Y-%m-%d')
report['质保时间'] = pd.to_datetime(report["质保时间"],errors='coerce').dt.strftime('%Y-%m-%d')
report[date_col] =report[date_col].astype(str)
report['终验收时间'] = ['' if i == '1990-01-01' else i
                              for i in report['终验收时间']]
report['实际出货日期'] = ['' if i == '1990-01-01' else i
                              for i in report['实际出货日期']]
report['订单签订日期'] = ['' if i == '1990-01-01' else i
                              for i in report['订单签订日期']]
report['预付款应付时间'] = ['' if i == '1990-01-01' else i
                              for i in report['预付款应付时间']]
report['发货款应付时间'] = ['' if i == '1990-01-01' else i
                              for i in report['发货款应付时间']]
report['验收款应付时间'] = ['' if i == '1990-01-01' else i
                              for i in report['验收款应付时间']]
report['质保款应付时间'] = ['' if i == '1990-01-01' else i
                              for i in report['质保款应付时间']]
report['质保时间'] = ['' if i == '1990-01-01' else i
                              for i in report['质保时间']]
report['开票日期'] = ['' if i == '1990-01-01' else i
                              for i in report['开票日期']]

#写入表格
#创建自定义excel
def writer_contents(sheet, array, start_row, start_col, format=None,
                    percent_format=None, percentlist=[]):
    start_col = 0
    for col in array:
        if percentlist and (start_col in percentlist):
            sheet.write_column(start_row, start_col, col, percent_format)
        else:
            sheet.write_column(start_row, start_col, col, format)
        start_col += 1


now_time = time.strftime("%Y-%m-%d-%H",time.localtime(time.time()))
book_name='应收款表单'+now_time
workbook = xlsxwriter.Workbook(book_name+'.xlsx', {'nan_inf_to_errors': True})
worksheet6 = workbook.add_worksheet('应收款汇总')
worksheet1 = workbook.add_worksheet('应收款明细')
worksheet7 = workbook.add_worksheet('客户欠款汇总')
worksheet2 = workbook.add_worksheet('预付款汇总')
worksheet3 = workbook.add_worksheet('发货款汇总')
worksheet4 = workbook.add_worksheet('验收款汇总')
worksheet5 = workbook.add_worksheet('质保款汇总')
title_format = workbook.add_format({'font_name': 'Arial',
                                    'font_size': 10,
                                    'font_color':'white',
                                    'bg_color':'#1F4E78',
                                    'bold': True,
                                    'bold': True,
                                    'align':'center',
                                    'valign':'vcenter',
                                    'border':1,
                                    'border_color':'white'
                                    })

title_format.set_align('vcenter')

# col_format = copy.deepcopy(title_format)
# col_format.set_font_size(10)
# col_format.set_bold(False)
# col_format.set_text_wrap(True)
col_format = workbook.add_format({'font_name': 'Arial',
                                    'font_size': 8,
                                    'font_color':'white',
                                    'bg_color':'#1F4E78',
                                    'text_wrap':True,
                                    'border':1,
                                    'border_color':'white',
                                    'align':'center',
                                    'valign':'vcenter'
                                    })

data_format = workbook.add_format({'font_name': 'Arial',
                                    'font_size': 10,
                                    'align':'left',
                                    'valign':'vcenter'
                                    })
data_format1 = workbook.add_format({'font_name': 'Arial',
                                    'font_size': 10,
                                    'align':'center',
                                    'valign':'vcenter'
                                    })
num_percent_data_format = workbook.add_format({'font_name':'Arial',
                                                'font_size': 10,
                                                'align':'center',
                                                'valign':'vcenter',
                                                'num_format':'0.00%'
                                                })
worksheet1.merge_range('A1:W1', '合同信息明细', title_format)
worksheet1.merge_range('X1:AD1', '立项信息明细', title_format)
worksheet1.merge_range('AE1:AI1', '出货信息明细', title_format)
worksheet1.merge_range('AJ1:AP1', '验收信息明细', title_format)
worksheet1.merge_range('AQ1:AR1', '回款信息明细', title_format)
worksheet1.merge_range('AS1:AU1', '开票信息明细', title_format)
worksheet1.merge_range('AW1:BB1', '预付款明细', title_format)
worksheet1.merge_range('BC1:BH1', '发货款明细', title_format)
worksheet1.merge_range('BI1:BN1', '验收款明细', title_format)
worksheet1.merge_range('BO1:BT1', '质保款明细', title_format)
worksheet1.write_row("A2", report.columns, col_format)
worksheet1.merge_range('AV1:AV2', '当前状态', title_format)
report_percent_col=['实际出货完成率', '实际验收完成率','回款完成率', '开票完成率']
percent_col_numlist = [report.columns.tolist().index(i) for i in report_percent_col]
writer_contents(sheet=worksheet1, array=report.T.values, start_row=2,
                            start_col=0,percent_format=num_percent_data_format,
                                percentlist=percent_col_numlist)

worksheet1.merge_range('BU1:BU2', '当前欠款总计', title_format)
worksheet1.merge_range('BV1:BV2', '当前还款结余', title_format)
#设置单元格宽
worksheet1.set_column('A:A', 5, data_format)
worksheet1.set_column('B:B', 5, data_format)
worksheet1.set_column('C:C', 7, data_format)
worksheet1.set_column('D:D', 10, data_format)
worksheet1.set_column('E:E', 12, data_format)
worksheet1.set_column('F:F', 7, data_format)
worksheet1.set_column('G:G', 9, data_format)
worksheet1.set_column('H:H', 10, data_format)
worksheet1.set_column('I:I', 6, data_format1)
worksheet1.set_column('J:J', 6, data_format1)
worksheet1.set_column('K:K', 6, data_format)
worksheet1.set_column('L:L', 6, data_format1)
worksheet1.set_column('M:M', 8, data_format1)
worksheet1.set_column('N:N', 6, data_format1)
worksheet1.set_column('O:O', 8, data_format1)
worksheet1.set_column('P:P', 12, data_format)
worksheet1.set_column('Q:Q', 7, data_format1)
worksheet1.set_column('R:R', 7, data_format1)
worksheet1.set_column('S:S', 7, data_format1)
worksheet1.set_column('T:T', 7, data_format1)
worksheet1.set_column('U:U', 7, data_format1)
worksheet1.set_column('V:V', 7, data_format1)
worksheet1.set_column('W:W', 10, data_format)
worksheet1.set_column('X:X', 7, data_format)
worksheet1.set_column('Y:Y', 12, data_format)
worksheet1.set_column('Z:Z', 9, data_format)
worksheet1.set_column('AA:AA', 7, data_format1)
worksheet1.set_column('AB:AB', 10, data_format)
worksheet1.set_column('AC:AC', 10, data_format)
worksheet1.set_column('AD:AD', 8, data_format)
worksheet1.set_column('AE:AE',8, data_format1)
worksheet1.set_column('AF:AF', 10, data_format1)
worksheet1.set_column('AG:AG', 5, data_format1)
worksheet1.set_column('AH:AH', 8, data_format1)
worksheet1.set_column('AI:AI', 10, data_format)
worksheet1.set_column('AJ:AJ', 5, data_format1)
worksheet1.set_column('AK:AK', 8, data_format1)
worksheet1.set_column('AL:AL', 8, data_format1)
worksheet1.set_column('AM:AM', 6, data_format1)
worksheet1.set_column('AN:AN', 8, data_format1)
worksheet1.set_column('AO:AO', 9, data_format)
worksheet1.set_column('AP:AP', 7,data_format1)
worksheet1.set_column('AQ:AQ', 8, data_format1)
worksheet1.set_column('AR:AR', 8, data_format1)
worksheet1.set_column('AS:AS', 9, data_format1)
worksheet1.set_column('AT:AT', 8, data_format1)
worksheet1.set_column('AU:AU', 10, data_format)
worksheet1.set_column('AV:AV', 12, data_format1)
worksheet1.set_column('AW:AW', 8, data_format1)
worksheet1.set_column('AX:AX', 8, data_format1)
worksheet1.set_column('AY:AY', 8, data_format1)
worksheet1.set_column('AZ:AZ', 10, data_format)
worksheet1.set_column('BA:BA', 8, data_format1)
worksheet1.set_column('BB:BB', 8, data_format1)
worksheet1.set_column('BC:BC', 8, data_format1)
worksheet1.set_column('BD:BD', 8, data_format1)
worksheet1.set_column('BE:BE', 8, data_format1)
worksheet1.set_column('BF:BF', 10, data_format)
worksheet1.set_column('BG:BG', 8, data_format1)
worksheet1.set_column('BH:BH', 8, data_format1)
worksheet1.set_column('BI:BI', 8, data_format1)
worksheet1.set_column('BJ:BJ', 8, data_format1)
worksheet1.set_column('BK:BK', 8, data_format1)
worksheet1.set_column('BL:BL', 10, data_format)
worksheet1.set_column('BM:BM', 8, data_format1)
worksheet1.set_column('BN:BN', 8, data_format1)
worksheet1.set_column('BO:BO', 8, data_format1)
worksheet1.set_column('BP:BP', 8, data_format1)
worksheet1.set_column('BQ:BQ', 8, data_format1)
worksheet1.set_column('BR:BR', 10, data_format)
worksheet1.set_column('BS:BS', 8, data_format1)
worksheet1.set_column('BT:BT', 10, data_format1)
worksheet1.set_column('BU:BU', 10, data_format1)
worksheet1.set_column('BV:BV', 10, data_format1)
List=np.array(report['序号']).tolist()
dic = {}
for i in List :
    if List.count(i)>1:
        dic[i] = List.count(i)
for i in dic.keys():
    #report[report['序号'] == i].index.tolist()[0]
    worksheet1.merge_range('%s:%s'%('A'+str(report[report['序号'] == i].index.tolist()[0]+3),'A'+str(report[report['序号'] == i].index.tolist()[0]+2+dic[i])), i, title_format)
    worksheet1.merge_range('%s:%s' % ('J' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'J' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])), report['订单数量'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('M' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'M' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['含税金额'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('N' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'N' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['USD'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)

    worksheet1.merge_range('%s:%s' % ('AA' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AA' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['立项数量'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
   # worksheet1.merge_range('%s:%s' % ('AE' + str(report[report['序号'] == i].index.tolist()[0] + 3),
    #                                  'AE' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
     #                      report['实际数量'][report[report['序号'] == i].index.tolist()[0]],
      #                     data_format1)
    worksheet1.merge_range('%s:%s' % ('AF' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AF' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['实际未税金额'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('AG' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AG' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['USD.1'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)


   # worksheet1.merge_range('%s:%s' % ('AJ' + str(report[report['序号'] == i].index.tolist()[0] + 3),
    #                                  'AJ' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
     #                      report['数量'][report[report['序号'] == i].index.tolist()[0]],
      #                     data_format1)
    worksheet1.merge_range('%s:%s' % ('AK' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AK' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['未税金额.1'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('AL' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AL' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['含税金额.1'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)

    worksheet1.merge_range('%s:%s' % ('AM' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AM' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['USD.2'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)

    worksheet1.merge_range('%s:%s' % ('AQ' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AQ' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['回款原币金额'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('AR' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AR' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['回款完成率'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('AS' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AS' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['开票总金额'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('AT' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AT' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['开票完成率'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('AW' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AW' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['预付款占比'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('AX' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AX' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['预付款应付'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('AY' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AY' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['预付款欠款'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('AZ' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'AZ' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['预付款应付时间'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('BA' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'BA' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['预付款账龄'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('BB' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'BB' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['预付款当前欠款'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)

    worksheet1.merge_range('%s:%s' % ('BU' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'BU' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['当前欠款总计'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
    worksheet1.merge_range('%s:%s' % ('BV' + str(report[report['序号'] == i].index.tolist()[0] + 3),
                                      'BV' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
                           report['当前实付结余'][report[report['序号'] == i].index.tolist()[0]],
                           data_format1)
  #  worksheet1.merge_range('%s:%s' % ('AZ' + str(report[report['序号'] == i].index.tolist()[0] + 3),
   #                                   'AZ' + str(report[report['序号'] == i].index.tolist()[0] + 2 + dic[i])),
    #                       report['预付款应付时间'][report[report['序号'] == i].index.tolist()[0]],
     #                      data_format1)
#
end = len(report) + 3
worksheet1.conditional_format('AH%s:AH%s'%(3,end), {'type': 'data_bar',
                                                        'bar_color': '#FABF8F',
                                                        'data_bar_2010':True,
                                                        'bar_solid':False})
worksheet1.conditional_format('AN%s:AN%s'%(3,end), {'type': 'data_bar',
                                                        'bar_color': '#FFFF00',
                                                        'data_bar_2010':True,
                                                        'bar_solid':False})
worksheet1.conditional_format('AR%s:AR%s'%(3,end), {'type': 'data_bar',
                                                        'bar_color': '#00FF00',
                                                        'data_bar_2010':True,
                                                        'bar_solid':False})
worksheet1.conditional_format('AT%s:AT%s'%(3,end), {'type': 'data_bar',
                                                        'bar_color': '#72F5FC',
                                                        'data_bar_2010':True,
                                                        'bar_solid':False})
#预付款
report_start.loc['Col'] =report_start.iloc[0:len(report_start),1:8].sum(axis=0)
report_start.loc['Col','公司简称']='汇总'
report_start=report_start.sort_values(by=['预付款应收'],axis=0,ascending=False).reset_index(drop=True)
tr=['公司简称','账款说明']
report_start[tr]=report_start[tr].fillna('')

worksheet2.write_row("A1", report_start.columns, title_format)
writer_contents(sheet=worksheet2, array=report_start.T.values, start_row=1,
                            start_col=0)
worksheet2.set_column('A:A', 8, data_format)
worksheet2.set_column('B:B', 14, data_format1)
worksheet2.set_column('C:C', 16, data_format1)
worksheet2.set_column('D:D', 16, data_format1)
worksheet2.set_column('E:E', 12, data_format1)
worksheet2.set_column('F:F', 12, data_format1)
worksheet2.set_column('G:G', 13, data_format)
worksheet2.set_column('H:H', 12, data_format1)
worksheet2.set_column('I:I', 80, data_format)
#发货款
report_outt.loc['Col'] =report_outt.iloc[0:len(report_outt),1:8].sum(axis=0)
report_outt.loc['Col','公司简称']='汇总'
report_outt=report_outt.sort_values(by=['发货款应收'],axis=0,ascending=False).reset_index(drop=True)
report_outt[tr]=report_outt[tr].fillna('')

worksheet3.write_row("A1", report_outt.columns, title_format)
writer_contents(sheet=worksheet3, array=report_outt.T.values, start_row=1,
                            start_col=0)
worksheet3.set_column('A:A', 8, data_format)
worksheet3.set_column('B:B', 14, data_format1)
worksheet3.set_column('C:C', 16, data_format1)
worksheet3.set_column('D:D', 16, data_format1)
worksheet3.set_column('E:E', 12, data_format1)
worksheet3.set_column('F:F', 12, data_format1)
worksheet3.set_column('G:G', 13, data_format)
worksheet3.set_column('H:H', 12, data_format1)
worksheet3.set_column('I:I', 80, data_format)
#验收款
report_rece.loc['Col'] =report_rece.iloc[0:len(report_rece),1:8].sum(axis=0)
report_rece.loc['Col','公司简称']='汇总'
report_rece=report_rece.sort_values(by=['验收款应收'],axis=0,ascending=False).reset_index(drop=True)
report_rece[tr]=report_rece[tr].fillna('')

worksheet4.write_row("A1", report_rece.columns, title_format)
writer_contents(sheet=worksheet4, array=report_rece.T.values, start_row=1,
                            start_col=0)
worksheet4.set_column('A:A', 8, data_format)
worksheet4.set_column('B:B', 14, data_format1)
worksheet4.set_column('C:C', 16, data_format1)
worksheet4.set_column('D:D', 16, data_format1)
worksheet4.set_column('E:E', 12, data_format1)
worksheet4.set_column('F:F', 12, data_format1)
worksheet4.set_column('G:G', 13, data_format)
worksheet4.set_column('H:H', 12, data_format1)
worksheet4.set_column('I:I', 80, data_format)
#质保款
report_pro.loc['Col'] =report_pro.iloc[0:len(report_pro),1:8].sum(axis=0)
report_pro.loc['Col','公司简称']='汇总'
report_pro=report_pro.sort_values(by=['质保款应收'],axis=0,ascending=False).reset_index(drop=True)
report_pro[tr]=report_pro[tr].fillna('')

worksheet5.write_row("A1", report_pro.columns, title_format)
writer_contents(sheet=worksheet5, array=report_pro.T.values, start_row=1,
                            start_col=0)
worksheet5.set_column('A:A', 8, data_format)
worksheet5.set_column('B:B', 14, data_format1)
worksheet5.set_column('C:C', 16, data_format1)
worksheet5.set_column('D:D', 16, data_format1)
worksheet5.set_column('E:E', 12, data_format1)
worksheet5.set_column('F:F', 12, data_format1)
worksheet5.set_column('G:G', 13, data_format)
worksheet5.set_column('H:H', 12, data_format1)
worksheet5.set_column('I:I', 80, data_format)


#汇总表
worksheet6.write_row("A1", report_to.columns, title_format)
writer_contents(sheet=worksheet6, array=report_to.T.values, start_row=1,
                            start_col=0)
worksheet6.set_column('A:A', 8, data_format)
worksheet6.set_column('B:B', 14, data_format1)
worksheet6.set_column('C:C', 16, data_format1)
worksheet6.set_column('D:D', 16, data_format1)
worksheet6.set_column('E:E', 16, data_format1)
worksheet6.set_column('F:F', 14, data_format1)
worksheet6.set_column('G:G', 13, data_format)
worksheet6.set_column('H:H', 12, data_format1)
#客户表

worksheet7.write_row("A1", report_cus.columns, title_format)
writer_contents(sheet=worksheet7, array=report_cus.T.values, start_row=1,
                            start_col=0)
worksheet7.set_column('A:A', 8, data_format)
worksheet7.set_column('B:B', 14, data_format1)
worksheet7.set_column('C:C', 16, data_format1)
worksheet7.set_column('D:D', 16, data_format1)
worksheet7.set_column('E:E', 16, data_format1)
worksheet7.set_column('F:F', 14, data_format1)
worksheet7.set_column('G:G', 13, data_format1)
worksheet7.set_column('H:H', 12, data_format1)
worksheet7.set_column('I:I', 12, data_format1)
worksheet7.set_column('J:J', 12, data_format1)
for i in range(1,len(report_cus)+1):
    if i%4==0:
        worksheet7.merge_range('%s:%s' % ('A' + str(i-2),  'A' + str(i+1)),
                           report_cus['公司简称'][i-4],
                           data_format1)
        worksheet7.merge_range('%s:%s' % ('J' + str(i - 2), 'J' + str(i + 1)),
                               report_cus['汇总'][i - 4],
                               data_format1)
workbook.close()
end_time = time.time()
print('执行时长:%d秒' % (end_time - start_time))