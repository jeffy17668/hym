import pandas as pd
import time
import xlsxwriter
import os
import shutil
from tkinter import filedialog
current_file_path = os.getcwd()
os.chdir(current_file_path)
import warnings
warnings.filterwarnings('ignore')
start_time = time.time()
#导入数据
filePath1 = '数据源'
file_name1 = os.listdir(filePath1)
for i in range(len(file_name1)):
    if str(file_name1[i]).count('~$') == 0:
        order=pd.read_excel(filePath1 + '/' + str(file_name1[i]) ,sheet_name='1、订单')
        out=pd.read_excel(filePath1 + '/' + str(file_name1[i]) ,sheet_name='2.出货')
        accept=pd.read_excel(filePath1 + '/' + str(file_name1[i]) ,sheet_name='3.验收')

#订单：处理订单表下数据格式，提取所需数据
default_date = '1990/01/01'
now_year = int(time.strftime("%Y",time.localtime(time.time())))
order["收到订单日期"]=order["收到订单日期"].fillna(default_date)
order["收到订单日期"] = pd.to_datetime(order["收到订单日期"],errors='coerce')
order["立项数量"]=order["立项数量"].fillna(0)
order["项目号"]=order["项目号"].fillna('空值')
order["大项目名称"]=order["大项目名称"].fillna("空值")
order.loc[order['产品线'].str.contains("激光A|激光B|自动化A|自动化B",na=False),'产品线'] = '大装配线'

report_order=order[(order["收到订单日期"] >=pd.Timestamp(now_year, 1, 1))&(order["产品线"].str.contains("配件|人力|改造")==False)&(order["项目类型"].str.contains("配件|人力|改造")==False)&(order["大项目名称"]!='空值')].reset_index(drop=True)
#订单：提取订单表相应字段
report_ord=report_order[['大项目名称','产品线','公司简称','项目号','业务员','项目类型','收到订单日期','立项数量','未税金额']]
report_ord.loc[report_ord['项目号'].str.contains("-R"),'立项数量'] = 0

#出货：处理到货表下数据格式，提取所需数据
out["实际出货日期"]=out["实际出货日期"].fillna(default_date)
out["实际出货日期"]= pd.to_datetime(out["实际出货日期"],errors='coerce')
out["实际数量"]=out["实际数量"].fillna(0)
out_str=['项目号','大项目名称','公司简称','产品线']
out["大项目名称"]=out["大项目名称"].fillna("空值")
out.loc[order['产品线'].str.contains("激光A|激光B|自动化A|自动化B",na=False),'产品线'] = '大装配线'
report_out=out[(out["实际出货日期"] >=pd.Timestamp(now_year, 1, 1))&(out["产品线"].str.contains("配件|人力|改造")==False)&(out["大项目名称"]!='空值')].reset_index(drop=True)

#验收：处理验收表下数据格式，提取所需数据
accept["系统录入日期"]=accept["系统录入日期"].fillna(default_date)
accept["系统录入日期"]= pd.to_datetime(accept["系统录入日期"],errors='coerce')
accept["数量"]=accept["数量"].fillna(0)
accept_str=['项目号','大项目名称','公司简称','产品线']
accept["大项目名称"]=accept["大项目名称"].fillna("空值")
accept.loc[order['产品线'].str.contains("激光A|激光B|自动化A|自动化B",na=False),'产品线'] = '大装配线'
report_accept=accept[(accept["系统录入日期"] >=pd.Timestamp(now_year, 1, 1))&(accept["产品线"].str.contains("配件|人力|改造")==False)&(accept["大项目名称"]!='空值')].reset_index(drop=True)

# 出货：获取订单表的收到订单日期
order1=order[['订单序号','收到订单日期','业务员']].drop_duplicates(subset=['订单序号'],keep='first').reset_index(drop=True)
report_out=pd.merge(report_out,order1,left_on='订单序号',right_on='订单序号',how='left')
report_outt=report_out[['大项目名称','产品线','公司简称','收到订单日期','项目号','业务员','实际出货日期','实际数量','实际未税金额']]
report_outt.loc[report_outt['项目号'].str.contains("-R"),'实际数量'] = 0

# 验收：获取订单表的收到订单日期
#order1=order[['订单序号','收到订单日期','业务员']].drop_duplicates(subset=['订单序号'],keep='first').reset_index(drop=True)
report_accept=pd.merge(report_accept,order1,left_on='订单序号',right_on='订单序号',how='left')
report_acceptt=report_accept[['大项目名称','产品线','公司简称','收到订单日期','项目号','业务员','系统录入日期','数量','未税金额']]
report_acceptt.loc[report_acceptt['项目号'].str.contains("-R"),'数量'] = 0

#订单：统一订单同一  大项目号和产品线下的业务员和公司简称
report_ord.sort_values(by=['大项目名称','产品线'],inplace=True)
for i in range(1,len(report_ord)):
    if report_ord.loc[i,'大项目名称']==report_ord.loc[i-1,'大项目名称'] and report_ord.loc[i,'产品线']==report_ord.loc[i-1,'产品线']:
        report_ord.loc[i,'公司简称']=report_ord.loc[i-1,'公司简称']
        report_ord.loc[i,'业务员']=report_ord.loc[i-1,'业务员']
#出货：统一出货  大项目号和产品线下的业务员和公司简称
report_outt.sort_values(by=['大项目名称','产品线'],inplace=True)
for i in range(1,len(report_outt)):
    if report_outt.loc[i,'大项目名称']==report_outt.loc[i-1,'大项目名称'] and report_outt.loc[i,'产品线']==report_outt.loc[i-1,'产品线']:
        report_outt.loc[i,'公司简称']=report_outt.loc[i-1,'公司简称']
        report_outt.loc[i,'业务员']=report_outt.loc[i-1,'业务员']
#验收：统一验收  大项目号和产品线下的业务员和公司简称
report_acceptt.sort_values(by=['大项目名称','产品线'],inplace=True)
for i in range(1,len(report_acceptt)):
    if report_acceptt.loc[i,'大项目名称']==report_acceptt.loc[i-1,'大项目名称'] and report_acceptt.loc[i,'产品线']==report_acceptt.loc[i-1,'产品线']:
        report_acceptt.loc[i,'公司简称']=report_acceptt.loc[i-1,'公司简称']
        report_acceptt.loc[i,'业务员']=report_acceptt.loc[i-1,'业务员']
#三表提取月份字段
report_ord['月份']=report_ord['收到订单日期'].dt.month
report_outt['月份']=report_outt['实际出货日期'].dt.month
report_outt['订单年份']=report_outt['收到订单日期'].dt.year
report_acceptt['月份']=report_acceptt['系统录入日期'].dt.month
report_acceptt['订单年份']=report_acceptt['收到订单日期'].dt.year

#订单：根据现在的月份创建字段并填充每行对应月份数据
now_mon = int(time.strftime("%m",time.localtime(time.time())))
mn_list=[]
for i in range(1,now_mon+1):
    #report_ord[str(i)+'月']=''
    mn_list.append(str(i)+'月')
    for j in range(len(report_ord)):
        if report_ord.loc[j,'月份']==i:
            report_ord.loc[j,str(i)+'月']=report_ord.loc[j,'未税金额']
        else:
            report_ord.loc[j,str(i)+'月']=0
cal_col=['立项数量','未税金额']+mn_list #待总计字段名
cal_col_new=[]
for i in cal_col:
    cal_col_new.append(i+"-总计")
report_orde=report_ord.groupby(['大项目名称','产品线','公司简称','业务员'])[cal_col].sum().add_suffix('-总计').reset_index()
report_orde.sort_values(by=['产品线'],inplace=True)
report_orde_group = report_orde.groupby(['产品线'])[cal_col_new].sum().add_suffix('').reset_index()
report_orde_group['产品线']=report_orde_group['产品线']+'龠总计'
l_en=len(report_orde)
le_n=len(report_orde.columns)
report_orde.loc['汇总'] =report_orde.iloc[0:l_en,4:le_n].sum(axis=0)
report_orde.loc['汇总','大项目名称']='统计'
report_orde.loc['汇总','产品线']='01111111111'

report_orde = pd.concat([report_orde, report_orde_group])
report_orde.sort_values(by=['产品线'],inplace=True)
report_orde=report_orde.fillna(' ')
report_orde['产品线']=report_orde['产品线'].replace('龠','',regex=True).astype(str)
report_orde['产品线']=report_orde['产品线'].replace('01111111111','',regex=True).astype(str)
cal_col_new.remove('立项数量-总计')
report_orde[cal_col_new]=report_orde[cal_col_new]/10000
report_orde[cal_col_new]=report_orde[cal_col_new].round(2)
#出货：根据现在的月份创建字段并填充每行对应月份数据
mn_list1=[]
for i in range(1,now_mon+1):
    #report_ord[str(i)+'月']=''
    mn_list1.append(str(i)+'月')
    for j in range(len(report_outt)):
        if report_outt.loc[j,'月份']==i:
            report_outt.loc[j,str(i)+'月']=report_outt.loc[j,'实际未税金额']
        else:
            report_outt.loc[j,str(i)+'月']=0
cal_col1=['实际数量','实际未税金额']+mn_list1 #待总计字段名
cal_col_new1=[]
for i in cal_col1:
    cal_col_new1.append(i+"-总计")
report_outtt=report_outt.groupby(['大项目名称','产品线','订单年份','公司简称','业务员'])[cal_col1].sum().add_suffix('-总计').reset_index()
report_outtt.sort_values(by=['产品线'],inplace=True)
report_outtt_group = report_outtt.groupby(['产品线'])[cal_col_new1].sum().add_suffix('').reset_index()
report_outtt_group['产品线']=report_outtt_group['产品线']+'龠总计'

len1=len(report_outtt)
len1_1=len(report_outtt.columns)
report_outtt.loc['汇总'] =report_outtt.iloc[0:len1,5:len1_1].sum(axis=0)
report_outtt.loc['汇总','大项目名称']='统计'
report_outtt.loc['汇总','产品线']='01111111111'

report_outtt = pd.concat([report_outtt, report_outtt_group])

report_outtt.sort_values(by=['产品线'],inplace=True)
report_outtt=report_outtt.fillna(' ')
report_outtt['产品线']=report_outtt['产品线'].replace('龠','',regex=True).astype(str)
report_outtt['产品线']=report_outtt['产品线'].replace('01111111111','',regex=True).astype(str)

cal_col_new1.remove('实际数量-总计')
report_outtt[cal_col_new1]=report_outtt[cal_col_new1]/10000
report_outtt[cal_col_new1]=report_outtt[cal_col_new1].round(2)
#验收：根据现在的月份创建字段并填充每行对应月份数据
mn_list2=[]
for i in range(1,now_mon+1):
    #report_ord[str(i)+'月']=''
    mn_list2.append(str(i)+'月')
    for j in range(len(report_acceptt)):
        if report_acceptt.loc[j,'月份']==i:
            report_acceptt.loc[j,str(i)+'月']=report_acceptt.loc[j,'未税金额']
        else:
            report_acceptt.loc[j,str(i)+'月']=0
cal_col2=['数量','未税金额']+mn_list2 #待总计字段名
cal_col_new2=[]
for i in cal_col2:
    cal_col_new2.append(i+"-总计")
report_accepttt=report_acceptt.groupby(['大项目名称','产品线','订单年份','公司简称','业务员'])[cal_col2].sum().add_suffix('-总计').reset_index()
report_accepttt.sort_values(by=['产品线'],inplace=True)
report_accepttt_group = report_accepttt.groupby(['产品线'])[cal_col_new2].sum().add_suffix('').reset_index()
report_accepttt_group['产品线']=report_accepttt_group['产品线']+'龠总计'

len2=len(report_accepttt)
len2_2=len(report_accepttt.columns)
report_accepttt.loc['汇总'] =report_accepttt.iloc[0:len2,5:len2_2].sum(axis=0)
report_accepttt.loc['汇总','大项目名称']='统计'
report_accepttt.loc['汇总','产品线']='01111111111'

report_accepttt = pd.concat([report_accepttt, report_accepttt_group])

report_accepttt.sort_values(by=['产品线'],inplace=True)
report_accepttt=report_accepttt.fillna(' ')
report_accepttt['产品线']=report_accepttt['产品线'].replace('龠','',regex=True).astype(str)
report_accepttt['产品线']=report_accepttt['产品线'].replace('01111111111','',regex=True).astype(str)
cal_col_new2.remove('数量-总计')
report_accepttt[cal_col_new2]=report_accepttt[cal_col_new2]/10000
report_accepttt[cal_col_new2]=report_accepttt[cal_col_new2].round(2)
#写入excel
def writer_contents(sheet, array, start_row, start_col, format=None,
                    percent_format=None, percentlist=[]):
    start_col = 0
    for col in array:
        if percentlist and (start_col in percentlist):
            sheet.write_column(start_row, start_col, col, percent_format)
        else:
            sheet.write_column(start_row, start_col, col, format)
        start_col += 1

now_time = time.strftime("%Y-%m-%d",time.localtime(time.time()))
book_name='滚动计划实绩效达成表'+now_time

workbook = xlsxwriter.Workbook(book_name+'.xlsx', {'nan_inf_to_errors': True})
worksheet1 = workbook.add_worksheet('订单')
worksheet2 = workbook.add_worksheet('出货')
worksheet3 = workbook.add_worksheet('验收')

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
statis_format2 = workbook.add_format({'font_name':'Arial',   #系列总计
                                                'font_size': 9,
                                                'align':'center',
                                                'valign':'vcenter',
                                                'bg_color':'#92CDDC'
                                                })
#订单
worksheet1.write_row("A1", report_orde.columns, title_format)
writer_contents(sheet=worksheet1, array=report_orde.T.values, start_row=1,
                            start_col=0)
row_count=len(report_orde)

for row_index in range(row_count):
        if '总计' in str(report_orde.iloc[row_index,1]):
            worksheet1.write_row(row_index+1,0,report_orde.iloc[row_index,:4].reset_index(drop=True).values,statis_format2)
worksheet1.set_column('A:A', 24,data_format)
worksheet1.set_column('B:B', 16,data_format)
worksheet1.set_column('C:C', 9,data_format)
worksheet1.set_column('D:D', 9,data_format)
worksheet1.set_column('E:E', 11,data_format1)
worksheet1.set_column('F:F', 11,data_format1)
worksheet1.set_column('G:R', 9,data_format1)

#出货
worksheet2.write_row("A1", report_outtt.columns, title_format)
writer_contents(sheet=worksheet2, array=report_outtt.T.values, start_row=1,
                            start_col=0)
row_count1=len(report_outtt)
for row_index in range(row_count1):
        if '总计' in str(report_outtt.iloc[row_index,1]):
            worksheet2.write_row(row_index+1,0,report_outtt.iloc[row_index,:4].reset_index(drop=True).values,statis_format2)
worksheet2.set_column('A:A', 24,data_format)
worksheet2.set_column('B:B', 16,data_format)
worksheet2.set_column('C:C', 9,data_format)
worksheet2.set_column('D:D', 9,data_format)
worksheet2.set_column('E:E', 9,data_format1)
worksheet2.set_column('F:F', 11,data_format1)
worksheet2.set_column('G:G', 11,data_format1)
worksheet2.set_column('H:R', 9,data_format1)

#验收
worksheet3.write_row("A1", report_accepttt.columns, title_format)
writer_contents(sheet=worksheet3, array=report_accepttt.T.values, start_row=1,
                            start_col=0)
row_count2=len(report_accepttt)
for row_index in range(row_count2):
        if '总计' in str(report_accepttt.iloc[row_index,1]):
            worksheet3.write_row(row_index+1,0,report_accepttt.iloc[row_index,:4].reset_index(drop=True).values,statis_format2)
worksheet3.set_column('A:A', 24,data_format)
worksheet3.set_column('B:B', 16,data_format)
worksheet3.set_column('C:C', 9,data_format)
worksheet3.set_column('D:D', 9,data_format)
worksheet3.set_column('E:E', 9,data_format1)
worksheet3.set_column('F:F', 11,data_format1)
worksheet3.set_column('G:G', 11,data_format1)
worksheet3.set_column('H:R', 9,data_format1)
workbook.close()
end_time = time.time()
print('执行时长:%d秒' % (end_time - start_time))