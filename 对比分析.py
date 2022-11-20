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

def write_color(book, sheet, data, fmt, col_num='AQ'):
    start = 3
    format_grey = book.add_format({'font_name': 'Arial',
                                        'font_size': 10,
                                        'bg_color':'#14F7E4'
                                        })
    format_grey.set_align('center')
    format_grey.set_align('vcenter')

    format_red = book.add_format({'font_name': 'Arial',
                                        'font_size': 10,
                                        'bg_color':'#F86470'
                                        })
    format_red.set_align('center')
    format_red.set_align('vcenter')

    for item in data:
        if item == '台账无':
            sheet.write(col_num + str(start), item, format_grey)
        elif item == '系统无':
            sheet.write(col_num + str(start), item, format_red)
        else:
            sheet.write(col_num + str(start), item, fmt)
        start += 1
def write_color1(book, sheet, data, fmt, col_num='AQ'):
    start = 2
    format_grey = book.add_format({'font_name': 'Arial',
                                        'font_size': 10,
                                        'bg_color':'#14F7E4'
                                        })
    format_grey.set_align('center')
    format_grey.set_align('vcenter')

    format_red = book.add_format({'font_name': 'Arial',
                                        'font_size': 10,
                                        'bg_color':'#F86470'
                                        })
    format_red.set_align('center')
    format_red.set_align('vcenter')

    for item in data:
        if item == '台账无':
            sheet.write(col_num + str(start), item, format_grey)
        elif item == '系统无':
            sheet.write(col_num + str(start), item, format_red)
        else:
            sheet.write(col_num + str(start), item, fmt)
        start += 1
#获取台账三个表数据
filePath = '数据源'
file_name = os.listdir(filePath)
for i in range(len(file_name)):
    if str(file_name[i]).count('~$') == 0:
        t_order=pd.read_excel(filePath + '/' + str(file_name[i]) ,sheet_name='1、订单')
        t_out=pd.read_excel(filePath + '/' + str(file_name[i]) ,sheet_name='2.出货')
        t_accept=pd.read_excel(filePath + '/' + str(file_name[i]) ,sheet_name='3.验收')

#生成系统订单表
filePath1 = r'对比数据源\订单'

file_name1 = os.listdir(filePath1)
for i in range(len(file_name1)):
    if str(file_name1[i]).count('~$') == 0:
        order=pd.read_excel(filePath1 + '/' + str(file_name1[i]),header=2 )
order['客户订购单号']=order['客户订购单号'].fillna('空值')
order1=order[(order['状态码说明']=='已审核')&(order['行状态']=='一般')&(order['客户简称'].str.contains("海目星",na=False)==False)&(order['客户订购单号'].str.contains("空值|无",na=False)==False)][['客户订购单号','客户简称']].drop_duplicates(subset=['客户订购单号','客户简称'],keep='first').reset_index(drop=True)
order1=order1.sort_values(by=['客户订购单号'],inplace=False).reset_index(drop=True)
order1['检查系统订购单号']=''
for i in range(0,len(order1)-1):
        if  order1.loc[i,'客户订购单号']==order1.loc[i+1,'客户订购单号']:
            order1.loc[i,'检查系统订购单号']='存在重复项'
            order1.loc[i+1,'检查系统订购单号']='存在重复项'
order1=order1.drop_duplicates(subset=['客户订购单号','客户简称'],keep='first').reset_index(drop=True)
#生成台账订单表
t_order['主订单号']=t_order['主订单号'].fillna('空值').astype(str)
t_order1=t_order[(t_order['系统录入单号']!='U8')&(t_order['系统录入单号']!='u8')&(t_order['主订单号']!='空值')][['主订单号','公司名称']].drop_duplicates(subset=['主订单号','公司名称'],keep='first').reset_index(drop=True)
t_order1['检查台账订单号']=''
for i in range(0,len(t_order1)-1):
        if  t_order1.loc[i,'主订单号']==t_order1.loc[i+1,'主订单号']:
            t_order1.loc[i,'检查台账订单号']='存在重复项'
            t_order1.loc[i+1,'检查台账订单号']='存在重复项'

#两边处理格式和对比数据
order1['属于系统']='系统'
order1['客户订购单号']=order1['客户订购单号'].replace(' ','',regex=True).astype(str)
order1['客户简称']=order1['客户简称'].replace(' ','',regex=True).astype(str)
order1['客户简称']=order1['客户简称'].str.strip()
order1['客户订购单号']=order1['客户订购单号'].str.strip()
order1['客户简称']=order1['客户简称'].replace('\(','（',regex=True).astype(str)
order1['客户简称']=order1['客户简称'].replace('\)','）',regex=True).astype(str)
t_order1= t_order1.rename(columns={'主订单号':'客户订购单号','公司名称':'客户简称'})
t_order1['属于台账']='台账'
t_order1['客户订购单号']=t_order1['客户订购单号'].replace(' ','',regex=True).astype(str)
t_order1['客户简称']=t_order1['客户简称'].replace(' ','',regex=True).astype(str)
t_order1['客户简称']=t_order1['客户简称'].replace('\(','（',regex=True).astype(str)
t_order1['客户简称']=t_order1['客户简称'].replace('\)','）',regex=True).astype(str)
t_order1['客户简称']=t_order1['客户简称'].str.strip()
t_order1['客户订购单号']=t_order1['客户订购单号'].str.strip()

t_order1=t_order1.drop_duplicates(subset=['客户订购单号','客户简称'],keep='first').reset_index(drop=True)
order1=order1.drop_duplicates(subset=['客户订购单号','客户简称'],keep='first').reset_index(drop=True)
order1_1=pd.merge(order1,t_order1,on=['客户订购单号','客户简称'],how='outer')

order1_1['对比分析'] = ''
for i in range(len(order1_1)):
  #  if order1_1.loc[i, '属于系统'] == '系统' and order1_1.loc[i, '属于台账'] == '台账':
   #     order1_1.loc[i, '对比结果'] = '一致'
    if order1_1.loc[i, '属于系统'] == '系统' and order1_1.loc[i, '属于台账'] != '台账':
        order1_1.loc[i, '对比分析'] = '台账无'
    if order1_1.loc[i, '属于系统'] != '系统' and order1_1.loc[i, '属于台账'] == '台账':
        order1_1.loc[i, '对比分析'] = '系统无'
order1_1[['客户简称','检查系统订购单号','属于系统','检查台账订单号','属于台账','对比分析']]=order1_1[['客户简称','检查系统订购单号','属于系统','检查台账订单号','属于台账','对比分析']].fillna('')
order1_1=order1_1.sort_values(by=['客户订购单号'],inplace=False).reset_index(drop=True)
for i in range(1,len(order1_1)):
    if order1_1.loc[i-1,'客户订购单号']==order1_1.loc[i,'客户订购单号']:
        order1_1.loc[i, '对比分析']='客户简称不一致'
        order1_1.loc[i-1, '对比分析'] = '客户简称不一致'


    #系统订单表

order2=order[(order['状态码说明']=='已审核')&(order['行状态']=='一般')&(order['客户简称'].str.contains("海目星",na=False)==False)&(order['客户订购单号'].str.contains("空值|无",na=False)==False)][['客户订购单号','项目编号','销售数量','币别说明','原币税前金额','原币含税金额']].reset_index(drop=True)
order2['客户订购单号'] = order2['客户订购单号'].str.strip()
order2[['客户订购单号','项目编号','币别说明']]=order2[['客户订购单号','项目编号','币别说明']].fillna('空值')
order2=order2.groupby(['客户订购单号','项目编号','币别说明'])['销售数量','原币税前金额','原币含税金额'].sum().add_suffix('-总计').reset_index()
order2=order2.sort_values(by=['客户订购单号','项目编号'],inplace=False).reset_index(drop=True)
order2['检查系统币别']=''
for i in range(0,len(order2)-1):
        if  order2.loc[i,'客户订购单号']==order2.loc[i+1,'客户订购单号'] and order2.loc[i,'币别说明']!=order2.loc[i+1,'币别说明'] :
            order2.loc[i,'检查系统币别']='存在两种'
            order2.loc[i+1,'检查系统币别']='存在两种'

#台账订单项目表·

t_order2=t_order[(t_order['系统录入单号']!='u8')&(t_order['主订单号']!='空值')][['主订单号','项目号','币别','立项数量','未税金额', '含税金额','USD']].reset_index(drop=True)
t_order2.loc[(t_order2['币别'].str.contains("USD|usd|Usd", na=False)), '含税金额'] = t_order2['USD']
t_order2['主订单号'] = t_order2['主订单号'].replace(' ', '', regex=True).astype(str)
t_order2['主订单号'] = t_order2['主订单号'].str.strip()
t_order2=t_order2.groupby(['主订单号','项目号','币别'])['立项数量','未税金额', '含税金额'].sum().add_suffix('-总计').reset_index()
t_order2=t_order2.sort_values(by=['主订单号','项目号'],inplace=False).reset_index(drop=True)
t_order2['检查台账币别']=''
for i in range(0,len(t_order2)-1):
        if  t_order2.loc[i,'主订单号']==t_order2.loc[i+1,'主订单号'] and t_order2.loc[i,'币别']!=t_order2.loc[i+1,'币别'] :
            t_order2.loc[i,'检查台账币别']='存在两种'
            t_order2.loc[i+1,'检查台账币别']='存在两种'

# 处理两边格式
order2['客户订购单号'] = order2['客户订购单号'].str.strip()
t_order2 = t_order2.rename(columns={'主订单号': '客户订购单号', '项目号': '项目编号'})
t_order2['客户订购单号'] = t_order2['客户订购单号'].replace(' ', '', regex=True).astype(str)
t_order2['客户订购单号'] = t_order2['客户订购单号'].str.strip()
order2_2 = pd.merge(order2, t_order2, on=['客户订购单号', '项目编号'], how='left')
order2_2.loc[(order2_2['币别'].str.contains("CNY|Cny|cny", na=False)), '币别'] = '人民币'
order2_2.loc[(order2_2['币别'].str.contains("USD|usd|Usd", na=False)), '币别'] = '美元'
order2_2.loc[(order2_2['币别说明'].str.contains("人", na=False)), '币别说明'] = '人民币'
order2_2.loc[(order2_2['币别说明'].str.contains("美", na=False)), '币别说明'] = '美元'
order2_2['未税金额-总计'] = order2_2['未税金额-总计'].fillna(987654321)
order2_2['币别'] = order2_2['币别'].fillna('空值')
order2_2['立项数量-总计'] = order2_2['立项数量-总计'].fillna(987654321)
order2_2['对比分析'] = ''
order2_2['对比数量'] = ''
order2_2['对比税前'] = ''
order2_2['对比税后'] = ''
order2_2['币别说明'] = order2_2['币别说明'].replace(' ', '', regex=True).astype(str)
order2_2['币别'] = order2_2['币别'].replace(' ', '', regex=True).astype(str)
order2_2['币别说明'] = order2_2['币别说明'].str.strip()
order2_2['币别'] = order2_2['币别'].str.strip()
for i in range(len(order2_2)):
    if order2_2.loc[i, '币别'] == '空值' and order2_2.loc[i, '未税金额-总计'] == 987654321 and order2_2.loc[
        i, '立项数量-总计'] == 987654321:
        order2_2.loc[i, '对比分析'] = '台账无'
    if order2_2.loc[i, '币别'] == order2_2.loc[i, '币别说明'] and abs(order2_2.loc[i, '销售数量-总计']-order2_2.loc[i, '立项数量-总计'])>1:
        order2_2.loc[i, '对比数量'] = '数量不一致'
    if order2_2.loc[i, '币别'] == order2_2.loc[i, '币别说明'] and abs(order2_2.loc[i, '原币税前金额-总计']-order2_2.loc[i, '未税金额-总计'])>1:
        order2_2.loc[i, '对比税前'] = '金额不一致'
    if order2_2.loc[i, '币别'] == order2_2.loc[i, '币别说明'] and abs(order2_2.loc[i, '原币含税金额-总计']-order2_2.loc[i, '含税金额-总计'])>1:
        order2_2.loc[i, '对比税后'] = '金额不一致'

order2_2.loc[(order2_2['币别'].str.contains("美元", na=False)), '对比税前'] = ''
order2_2.loc[(order2_2['对比分析'].str.contains("台账", na=False)), '对比税前'] = ''
order2_2.loc[(order2_2['对比分析'].str.contains("台账", na=False)), '对比税后'] = ''
order2_2.loc[(order2_2['对比分析'].str.contains("台账", na=False)), '对比数量'] = ''
order2_2['币别'] = order2_2['币别'].replace('空值', '', regex=True).astype(str)
order2_2['立项数量-总计'] = order2_2['立项数量-总计'].replace(987654321, '', regex=True)
order2_2['未税金额-总计'] = order2_2['未税金额-总计'].replace(987654321, '', regex=True)
order2_2[['客户订购单号', '项目编号', '币别说明']] = order2_2[['客户订购单号', '项目编号', '币别说明']].replace('空值', '', regex=True)
order2_2[['含税金额-总计', '检查台账币别']] = order2_2[['含税金额-总计', '检查台账币别']].fillna('')
order2_2=order2_2.sort_values(by=['客户订购单号'],inplace=False).reset_index(drop=True)
#系统订单表3·
order3=order[(order['状态码说明']=='已审核')&(order['行状态']=='一般')&(order['客户简称'].str.contains("海目星",na=False)==False)&(order['客户订购单号'].str.contains("空值|无",na=False)==False)][['客户订购单号','币别说明','原币税前金额','原币含税金额']].reset_index(drop=True)
order3[['客户订购单号','币别说明']]=order3[['客户订购单号','币别说明']].fillna('空值')
order3['客户订购单号'] = order3['客户订购单号'].str.strip()
order3=order3.groupby(['客户订购单号','币别说明'])['原币税前金额','原币含税金额'].sum().add_suffix('-总计').reset_index()
order3=order3.sort_values(by=['客户订购单号'],inplace=False).reset_index(drop=True)
#台账订单表3
t_order3=t_order[(t_order['系统录入单号']!='u8')&(t_order['主订单号']!='空值')][['主订单号','币别','未税金额', '含税金额','USD']].reset_index(drop=True)
t_order3.loc[(t_order3['币别'].str.contains("USD|usd|Usd", na=False)), '含税金额'] = t_order3['USD']
t_order3['主订单号'] = t_order3['主订单号'].replace(' ', '', regex=True).astype(str)
t_order3['主订单号'] = t_order3['主订单号'].str.strip()
t_order3=t_order3.groupby(['主订单号','币别'])['未税金额', '含税金额'].sum().add_suffix('-总计').reset_index()
t_order3=t_order3.sort_values(by=['主订单号'],inplace=False).reset_index(drop=True)

# 处理两边格式
order3['客户订购单号'] = order3['客户订购单号'].str.strip()
t_order3 = t_order3.rename(columns={'主订单号': '客户订购单号'})
t_order3['客户订购单号'] = t_order3['客户订购单号'].replace(' ', '', regex=True).astype(str)
t_order3['客户订购单号'] = t_order3['客户订购单号'].str.strip()
order3_3 = pd.merge(order3, t_order3, on=['客户订购单号'], how='left')
order3_3.loc[(order3_3['币别'].str.contains("CNY|Cny|cny", na=False)), '币别'] = '人民币'
order3_3.loc[(order3_3['币别'].str.contains("USD|usd|Usd", na=False)), '币别'] = '美元'
order3_3.loc[(order3_3['币别说明'].str.contains("人", na=False)), '币别说明'] = '人民币'
order3_3.loc[(order3_3['币别说明'].str.contains("美", na=False)), '币别说明'] = '美元'

order3_3['未税金额-总计'] = order3_3['未税金额-总计'].fillna(987654321)
order3_3['币别'] = order3_3['币别'].fillna('空值')

order3_3['对比分析'] = ''

order3_3['对比税前'] = ''
order3_3['对比税后'] = ''
order3_3['币别说明'] = order3_3['币别说明'].replace(' ', '', regex=True).astype(str)
order3_3['币别'] = order3_3['币别'].replace(' ', '', regex=True).astype(str)
order3_3['币别说明'] = order3_3['币别说明'].str.strip()
order3_3['币别'] = order3_3['币别'].str.strip()
for i in range(len(order3_3)):
    if order3_3.loc[i, '币别'] == '空值' and order3_3.loc[i, '未税金额-总计'] == 987654321 :
        order3_3.loc[i, '对比分析'] = '台账无'

    if order3_3.loc[i, '币别'] == order3_3.loc[i, '币别说明'] and abs(order3_3.loc[i, '原币税前金额-总计']-order3_3.loc[i, '未税金额-总计'])>1:
        order3_3.loc[i, '对比税前'] = '金额不一致'
    if order3_3.loc[i, '币别'] == order3_3.loc[i, '币别说明'] and abs(order3_3.loc[i, '原币含税金额-总计']-order3_3.loc[i, '含税金额-总计'])>1:
        order3_3.loc[i, '对比税后'] = '金额不一致'

order3_3.loc[(order3_3['币别'].str.contains("美元", na=False)), '对比税前'] = ''
order3_3.loc[(order3_3['对比分析'].str.contains("台账", na=False)), '对比税前'] = ''
order3_3.loc[(order3_3['对比分析'].str.contains("台账", na=False)), '对比税后'] = ''

order3_3['币别'] = order3_3['币别'].replace('空值', '', regex=True).astype(str)

order3_3['未税金额-总计'] = order3_3['未税金额-总计'].replace(987654321, '', regex=True)
order3_3[['客户订购单号',  '币别说明']] = order3_3[['客户订购单号',  '币别说明']].replace('空值', '', regex=True)
order3_3[['含税金额-总计']] = order3_3[['含税金额-总计']].fillna('')
order3_3=order3_3.sort_values(by=['客户订购单号'],inplace=False).reset_index(drop=True)
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
book_name='对比分析表'+now_time

workbook = xlsxwriter.Workbook(book_name+'.xlsx', {'nan_inf_to_errors': True})
worksheet1 = workbook.add_worksheet('订单-客户')
worksheet2 = workbook.add_worksheet('订单-项目号')
worksheet3 = workbook.add_worksheet('订单-合同号')

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

#订单1
worksheet1.write_row("A1", order1_1.columns, title_format)
writer_contents(sheet=worksheet1, array=order1_1.T.values, start_row=1,
                            start_col=0)
write_color1(book=workbook, sheet=worksheet1, data=order1_1['对比分析'],
                        fmt=data_format1, col_num='G')

worksheet1.set_column('A:A', 16,data_format)
worksheet1.set_column('B:B', 15,data_format)
worksheet1.set_column('C:C', 9,data_format1)
worksheet1.set_column('D:D', 9,data_format1)
worksheet1.set_column('E:E', 11,data_format1)
worksheet1.set_column('F:F', 11,data_format1)
worksheet1.set_column('G:G', 12,data_format1)
worksheet1.set_column('H:R', 9,data_format1)

#订单2
worksheet2.write_row("A2", order2_2.columns, title_format)
writer_contents(sheet=worksheet2, array=order2_2.T.values, start_row=2,
                            start_col=0)
write_color(book=workbook, sheet=worksheet2, data=order2_2['对比分析'],
                        fmt=data_format1, col_num='M')
worksheet2.merge_range('A1:G1', '系统信息', title_format)
worksheet2.merge_range('H1:L1', '台账信息', title_format)
worksheet2.merge_range('M1:P1', '对比结果', title_format)
worksheet2.set_column('A:A', 16,data_format)
worksheet2.set_column('B:B', 15,data_format)
worksheet2.set_column('C:C', 9,data_format1)
worksheet2.set_column('D:D', 10,data_format1)
worksheet2.set_column('E:E', 12,data_format1)
worksheet2.set_column('F:F', 12,data_format1)
worksheet2.set_column('G:L', 12,data_format1)
worksheet2.set_column('L:P', 10,data_format1)
#订单3

worksheet3.write_row("A2", order3_3.columns, title_format)
writer_contents(sheet=worksheet3, array=order3_3.T.values, start_row=2,
                            start_col=0)
write_color(book=workbook, sheet=worksheet3, data=order3_3['对比分析'],
                        fmt=data_format1, col_num='H')
worksheet3.merge_range('A1:D1', '系统信息', title_format)
worksheet3.merge_range('E1:G1', '台账信息', title_format)
worksheet3.merge_range('H1:J1', '对比结果', title_format)
worksheet3.set_column('A:A', 16,data_format)
worksheet3.set_column('B:B', 10,data_format)
worksheet3.set_column('C:C', 9,data_format1)
worksheet3.set_column('D:D', 9,data_format1)
worksheet3.set_column('E:E', 11,data_format1)
worksheet3.set_column('F:F', 13,data_format1)
worksheet3.set_column('G:R', 9,data_format1)

#出货
filePath2 = r'对比数据源\出货'
file_name2 = os.listdir(filePath2)
if os.listdir(filePath2):
    for i in range(len(file_name2)):
        if str(file_name2[i]).count('~$') == 0:
            out=pd.read_excel(filePath2 + '/' + str(file_name2[i]),header=2 )
    out['客户订购单号']=out['客户订购单号'].fillna('空值')
    out1=out[(out['状态.1'].str.contains("已审核|已过账",na=False)==True)&(out['客户简称'].str.contains("海目星",na=False)==False)&(out['客户订购单号'].str.contains("空值|无")==False)][['客户订购单号','项目编号','出货数量','币种说明','原币税前金额','原币含税金额']].reset_index(drop=True)
    out1=out1.sort_values(by=['客户订购单号'],inplace=False).reset_index(drop=True)
    out1[['客户订购单号','项目编号','币种说明']]=out1[['客户订购单号','项目编号','币种说明']].fillna('空值')
    out1['客户订购单号'] = out1['客户订购单号'].str.strip()
    out1=out1.groupby(['客户订购单号','项目编号','币种说明'])['出货数量','原币税前金额','原币含税金额'].sum().add_suffix('-总计').reset_index()
    out1=out1.sort_values(by=['客户订购单号','项目编号'],inplace=False).reset_index(drop=True)

    #生成台账出货表
    t_out1=t_out[(t_out['系统录入单号'].str.contains("U8|u8",na=False)==False)&(t_out['主订单号']!='空值')][['主订单号','项目号','币别','实际数量','实际未税金额', '实际含税金额','USD']].reset_index(drop=True)
    t_out1.loc[(t_out1['币别'].str.contains("USD|usd|Usd", na=False)), '实际含税金额'] = t_out1['USD']
    t_out1['主订单号'] = t_out1['主订单号'].replace(' ', '', regex=True).astype(str)
    t_out1['主订单号'] = t_out1['主订单号'].str.strip()
    t_out1=t_out1.groupby(['主订单号','项目号','币别'])['实际数量','实际未税金额', '实际含税金额'].sum().add_suffix('-总计').reset_index()
    t_out1=t_out1.sort_values(by=['主订单号','项目号'],inplace=False).reset_index(drop=True)
    # 处理两边格式
    out1['属于系统']='系统'
    t_out1['属于台账']='台账'
    out1['客户订购单号'] = out1['客户订购单号'].str.strip()
    t_out1 = t_out1.rename(columns={'主订单号': '客户订购单号', '项目号': '项目编号'})
    t_out1['客户订购单号'] = t_out1['客户订购单号'].replace(' ', '', regex=True).astype(str)
    t_out1['客户订购单号'] = t_out1['客户订购单号'].str.strip()

    out1_1 = pd.merge(out1, t_out1, on=['客户订购单号', '项目编号'], how='outer')
    out1_1.loc[(out1_1['币别'].str.contains("CNY|Cny|cny", na=False)), '币别'] = '人民币'
    out1_1.loc[(out1_1['币别'].str.contains("USD|usd|Usd", na=False)), '币别'] = '美元'
    out1_1.loc[(out1_1['币种说明'].str.contains("人", na=False)), '币种说明'] = '人民币'
    out1_1.loc[(out1_1['币种说明'].str.contains("美", na=False)), '币种说明'] = '美元'
    #out1_1['未税金额-总计'] = order3_3['未税金额-总计'].fillna(987654321)
    out1_1['属于系统'] = out1_1['属于系统'].fillna('空值')
    out1_1['属于台账'] = out1_1['属于台账'].fillna('空值')
    out1_1['对比分析']=''
    out1_1['对比数量'] = ''
    out1_1['对比税前'] = ''
    out1_1['对比税后'] = ''
    for i in range(len(out1_1)):
        if out1_1.loc[i,'属于台账']=='空值':
            out1_1.loc[i, '对比分析']='台账无'
        if out1_1.loc[i,'属于系统']=='空值':
            out1_1.loc[i, '对比分析']='系统无'
        if '无' not in str(out1_1.loc[i, '对比分析']):
            if abs(out1_1.loc[i, '实际数量-总计'] - out1_1.loc[i, '出货数量-总计']) > 1:
                out1_1.loc[i, '对比数量']='数量不一致'
            if abs(out1_1.loc[i, '原币税前金额-总计'] - out1_1.loc[i, '实际未税金额-总计']) > 1:
                out1_1.loc[i, '对比税前']='金额不一致'
            if abs(out1_1.loc[i, '原币含税金额-总计'] - out1_1.loc[i, '实际含税金额-总计']) > 1:
                out1_1.loc[i, '对比税后']='金额不一致'
    out1_1.loc[(out1_1['币别'].str.contains("美元", na=False)), '对比税前'] = ''
    out1_1.loc[(out1_1['币种说明'].str.contains("美元", na=False)), '对比税前'] = ''
    out1_1.loc[(out1_1['对比分析'].str.contains("无", na=False)), '对比税前'] = ''
    out1_1.loc[(out1_1['对比分析'].str.contains("无", na=False)), '对比税后'] = ''

    out1_1['属于台账'] = out1_1['属于台账'].replace('空值', '', regex=True).astype(str)
    out1_1['属于系统'] = out1_1['属于系统'].replace('空值', '', regex=True).astype(str)
    strl=['客户订购单号','项目编号','出货数量-总计','币种说明','原币税前金额-总计','原币含税金额-总计','币别','实际数量-总计','实际未税金额-总计', '实际含税金额-总计','对比分析','对比数量','对比税前','对比税后']
    #out1_1['未税金额-总计'] = out1_1['未税金额-总计'].replace(987654321, '', regex=True)
    #out1_1[['客户订购单号', '币别说明']] = out1_1[['客户订购单号', '币别说明']].replace('空值', '', regex=True)
    out1_1[strl] = out1_1[strl].fillna('')
    out1_1 = out1_1.sort_values(by=['客户订购单号'], inplace=False).reset_index(drop=True)


    worksheet4 = workbook.add_worksheet('出货')
    worksheet4.write_row("A2", out1_1.columns, title_format)
    writer_contents(sheet=worksheet4, array=out1_1.T.values, start_row=2, start_col=0)
    write_color(book=workbook, sheet=worksheet4, data=out1_1['对比分析'],
                fmt=data_format1, col_num='M')
    worksheet4.merge_range('A1:B1', '合同信息', title_format)
    worksheet4.merge_range('C1:G1', '系统信息', title_format)
    worksheet4.merge_range('H1:L1', '台账信息', title_format)
    worksheet4.merge_range('M1:P1', '对比结果', title_format)
    worksheet4.set_column('A:A', 16, data_format)
    worksheet4.set_column('B:B', 14, data_format)

    worksheet4.set_column('D:F', 15, data_format1)
    worksheet4.set_column('G:H', 10, data_format1)
    worksheet4.set_column('I:K', 15, data_format1)
    worksheet4.set_column('L:P', 10, data_format1)

#验收
filePath2 = r'对比数据源\验收'
file_name2 = os.listdir(filePath2)
if os.listdir(filePath2):
    #filenames = os.listdir(dir)
    index = 0
    dfs = []
    for name in file_name2:
        print(index)
        dfs.append(pd.read_excel(os.path.join(filePath2, name)))
        index += 1  # 为了查看合并到第几个表格了
    accept = pd.concat(dfs)
    accept=accept.reset_index(drop=True)
    accept['客户订购单号'] = accept['客户订购单号'].fillna('空值')
    accept1 = accept[(accept['交易对象名称'].str.contains("海目星", na=False) == False) & (
                                 accept['客户订购单号'].str.contains("空值|无", na=False) == False)][
        ['客户订购单号', '项目编号', '出货数量', '币种说明', '出货税前原币金额', '出货含税原币金额']].reset_index(drop=True)
    accept1 = accept1.sort_values(by=['客户订购单号'], inplace=False).reset_index(drop=True)
    accept1[['客户订购单号', '项目编号', '币种说明']] = accept1[['客户订购单号', '项目编号', '币种说明']].fillna('空值')

    accept1['客户订购单号'] = accept1['客户订购单号'].str.strip()
    accept1 = accept1.groupby(['客户订购单号', '项目编号', '币种说明'])['出货数量', '出货税前原币金额', '出货含税原币金额'].sum().add_suffix(
        '-总计').reset_index()
    accept1 = accept1.sort_values(by=['客户订购单号', '项目编号'], inplace=False).reset_index(drop=True)

    # 生成台账验收表
    t_accept1 = t_accept[(t_accept['系统录入单号'].str.contains("U8|u8",na=False)==False) & (t_accept['主订单号'] != '空值')][
        ['主订单号', '项目号', '币别', '数量', '未税金额', '含税金额','USD']].reset_index(drop=True)
    t_accept1.loc[(t_accept1['币别'].str.contains("USD|usd|Usd", na=False)), '含税金额'] = t_accept1['USD']
    t_accept1['主订单号'] = t_accept1['主订单号'].replace(' ', '', regex=True).astype(str)
    t_accept1['主订单号'] = t_accept1['主订单号'].str.strip()
    t_accept1 = t_accept1.groupby(['主订单号', '项目号', '币别'])['数量', '未税金额', '含税金额'].sum().add_suffix(
        '-总计').reset_index()
    t_accept1 = t_accept1.sort_values(by=['主订单号', '项目号'], inplace=False).reset_index(drop=True)
    # 处理两边格式
    accept1['属于系统'] = '系统'
    t_accept1['属于台账'] = '台账'
    accept1['客户订购单号'] = accept1['客户订购单号'].str.strip()
    #accept1 = accept1.rename(columns={'主订单号': '客户订购单号', '项目号': '项目编号'})
    t_accept1 = t_accept1.rename(columns={'主订单号': '客户订购单号', '项目号': '项目编号'})
    t_accept1['客户订购单号'] = t_accept1['客户订购单号'].replace(' ', '', regex=True).astype(str)
    t_accept1['客户订购单号'] = t_accept1['客户订购单号'].str.strip()

    accept1_1 = pd.merge(accept1, t_accept1, on=['客户订购单号', '项目编号'], how='outer')
    accept1_1.loc[(accept1_1['币别'].str.contains("CNY|Cny|cny", na=False)), '币别'] = '人民币'
    accept1_1.loc[(accept1_1['币别'].str.contains("USD|usd|Usd", na=False)), '币别'] = '美元'
    accept1_1.loc[(accept1_1['币种说明'].str.contains("人", na=False)), '币种说明'] = '人民币'
    accept1_1.loc[(accept1_1['币种说明'].str.contains("美", na=False)), '币种说明'] = '美元'
    # accept1_1['未税金额-总计'] = order3_3['未税金额-总计'].fillna(987654321)
    accept1_1['属于系统'] = accept1_1['属于系统'].fillna('空值')
    accept1_1['属于台账'] = accept1_1['属于台账'].fillna('空值')
    accept1_1['对比分析'] = ''
    accept1_1['对比数量'] = ''
    accept1_1['对比税前'] = ''
    accept1_1['对比税后'] = ''
    for i in range(len(accept1_1)):
        if accept1_1.loc[i, '属于台账'] == '空值':
            accept1_1.loc[i, '对比分析'] = '台账无'
        if accept1_1.loc[i, '属于系统'] == '空值':
            accept1_1.loc[i, '对比分析'] = '系统无'
        if '无' not in str(accept1_1.loc[i, '对比分析']):
            if abs(accept1_1.loc[i, '数量-总计'] - accept1_1.loc[i, '出货数量-总计']) > 1:
                accept1_1.loc[i, '对比数量'] = '数量不一致'
            if abs(accept1_1.loc[i, '出货税前原币金额-总计'] - accept1_1.loc[i, '未税金额-总计']) > 1:
                accept1_1.loc[i, '对比税前'] = '金额不一致'
            if abs(accept1_1.loc[i, '出货含税原币金额-总计'] - accept1_1.loc[i, '含税金额-总计']) > 1:
                accept1_1.loc[i, '对比税后'] = '金额不一致'
    accept1_1.loc[(accept1_1['币别'].str.contains("美元", na=False)), '对比税前'] = ''
    accept1_1.loc[(accept1_1['币种说明'].str.contains("美元", na=False)), '对比税前'] = ''
    accept1_1.loc[(accept1_1['对比分析'].str.contains("无", na=False)), '对比税前'] = ''
    accept1_1.loc[(accept1_1['对比分析'].str.contains("无", na=False)), '对比税后'] = ''

    accept1_1['属于台账'] = accept1_1['属于台账'].replace('空值', '', regex=True).astype(str)
    accept1_1['属于系统'] = accept1_1['属于系统'].replace('空值', '', regex=True).astype(str)
    strl1 = ['客户订购单号', '项目编号', '出货数量-总计', '币种说明', '出货税前原币金额-总计', '出货含税原币金额-总计', '币别', '数量-总计', '未税金额-总计', '含税金额-总计',
            '对比分析', '对比数量', '对比税前', '对比税后']
    # accept1_1['未税金额-总计'] = accept1_1['未税金额-总计'].replace(987654321, '', regex=True)
    # accept1_1[['客户订购单号', '币别说明']] = accept1_1[['客户订购单号', '币别说明']].replace('空值', '', regex=True)
    accept1_1[strl1] = accept1_1[strl1].fillna('')
    accept1_1 = accept1_1.sort_values(by=['客户订购单号'], inplace=False).reset_index(drop=True)

    worksheet5 = workbook.add_worksheet('验收')
    worksheet5.write_row("A2", accept1_1.columns, title_format)
    writer_contents(sheet=worksheet5, array=accept1_1.T.values, start_row=2, start_col=0)
    write_color(book=workbook, sheet=worksheet5, data=accept1_1['对比分析'],
                fmt=data_format1, col_num='M')
    worksheet5.merge_range('A1:B1', '合同信息', title_format)
    worksheet5.merge_range('C1:G1', '系统信息', title_format)
    worksheet5.merge_range('H1:L1', '台账信息', title_format)
    worksheet5.merge_range('M1:P1', '对比结果', title_format)
    worksheet5.set_column('A:A', 16, data_format)
    worksheet5.set_column('B:B', 14, data_format)
    worksheet5.set_column('D:F', 15, data_format1)
    worksheet5.set_column('G:H', 10, data_format1)
    worksheet5.set_column('I:K', 15, data_format1)
    worksheet5.set_column('L:P', 10, data_format1)

workbook.close()
end_time = time.time()
print('执行时长:%d秒' % (end_time - start_time))