import tkinter as tk
from tkinter import ttk  #下拉框
import tkinter.messagebox #弹出框
from tkinter import *
#from xlutils.copy import copy
from PIL import ImageTk
from tkinter import filedialog
import smtplib
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.application import MIMEApplication
import os
#使用MIMEMUltipart添加附件
from email.mime.multipart import MIMEMultipart
from email.header import Header
import traceback
# 创建窗口
window = tk.Tk()
window.title('邮件发送')  # 窗口的标题
window.geometry('500x800')  # 窗口的大小
window.iconbitmap('软件附带文件\头像.ico')
frame=tk.Canvas(window,width=500,height=800,background='silver',scrollregion=(0,0,500,3800))
roll=Scrollbar(window,orient='vertical',command=frame.yview)
#frame.pack(fill="both",side='right')
frame['yscrollcommand']=roll.set
roll.pack(side=RIGHT, fill=Y)
frame.pack(side=TOP, fill=Y, expand=True)

image_file = ImageTk.PhotoImage(file=r'软件附带文件\背景.jpg')
#发件人
send = tk.Label(frame,
             text='发件人',  # 标签的文字
             bg='darkred',  # 标签背景颜色
             font=('华文行楷', 15),  # 字体和字体大小
             width=10, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='white')
frame.create_window(52,25,window=send)
send_ma= tk.StringVar(value='zhengfei@hymson.com')
send_man=tk.Entry(frame,show=None,width=36,bd=4,cursor='cross',textvariable=send_ma)
frame.create_window(250,25,window=send_man)
#密码
pass_= tk.Label(frame,
             text='密码',  # 标签的文字
             bg='darkred',  # 标签背景颜色
             font=('华文行楷', 15),  # 字体和字体大小
             width=10, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='white')
frame.create_window(52,85,window=pass_)
passwor= tk.StringVar(value='Zf@340825')
password=tk.Entry(frame,show=None,width=22,bd=4,cursor='cross',textvariable=passwor)
frame.create_window(200,85,window=password)

#收件人
rece= tk.Label(frame,
             text='收件人',  # 标签的文字
             bg='darkred',  # 标签背景颜色
             font=('华文行楷', 15),  # 字体和字体大小
             width=10, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='white')
frame.create_window(52,145,window=rece)
receive= tk.StringVar(value='liyan@hymson.com,zhengfei@hymson.com')
receivior=tk.Entry(frame,show=None,width=36,bd=4,cursor='cross',textvariable=receive)
frame.create_window(250,145,window=receivior)

#发件人昵称
send_na= tk.Label(frame,
             text='发件人昵称',  # 标签的文字
             bg='darkred',  # 标签背景颜色
             font=('华文行楷', 15),  # 字体和字体大小
             width=10, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='white')
frame.create_window(52,205,window=send_na)
send_nam= tk.StringVar(value='李燕')
send_name=tk.Entry(frame,show=None,width=15,bd=4,cursor='cross',textvariable=send_nam)
frame.create_window(180,205,window=send_name)

#收件人昵称
rece_na= tk.Label(frame,
             text='收件人昵称',  # 标签的文字
             bg='darkred',  # 标签背景颜色
             font=('华文行楷', 15),  # 字体和字体大小
             width=10, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='white')
frame.create_window(52,265,window=rece_na)
rece_nam= tk.StringVar(value='相关同事')
rece_name=tk.Entry(frame,show=None,width=15,bd=4,cursor='cross',textvariable=rece_nam)
frame.create_window(180,265,window=rece_name)

#邮件主题
titl= tk.Label(frame,
             text='邮件主题',  # 标签的文字
             bg='darkred',  # 标签背景颜色
             font=('华文行楷', 15),  # 字体和字体大小
             width=10, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='white')
frame.create_window(52,325,window=titl)
titl1= tk.StringVar(value='长期未出货及长期未验收明细自动邮件')
title=tk.Entry(frame,show=None,width=36,bd=4,cursor='cross',textvariable=titl1)
frame.create_window(250,325,window=title)

#邮件内容
te= tk.Label(frame,
             text='邮件内容',  # 标签的文字
             bg='darkred',  # 标签背景颜色
             font=('华文行楷', 15),  # 字体和字体大小
             width=10, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='white')
frame.create_window(52,385,window=te)
tex= tk.StringVar(value='长期未出货及长期未验收明细自动邮件相关')
text=tk.Entry(frame,show=None,width=36,bd=4,cursor='cross',textvariable=tex)
frame.create_window(250,385,window=text)
#附件地址
def upload_file():
    selectFile = tk.filedialog.askopenfilename()  # askopenfilename 1次上传1个；askopenfilenames1次上传多个
    entry2.insert(0, selectFile)
entry2 = tk.Entry(frame, width=36,bd=4,textvariable= tk.StringVar())
frame.create_window(250,445,window=entry2)
btn= tk.Button(frame, text='上传附件', bg='darkred',  # 标签背景颜色
             font=('华文行楷', 15),  # 字体和字体大小width=10, height=1,  # 标签长宽(以字符长度计算)
             padx=2,pady=4,anchor='s',fg='white',bd=4, command=upload_file)
frame.create_window(52,445,window=btn)

my_sender = send_man.get()
# user登录邮箱的用户名，password登录邮箱的密码（授权码，即客户端密码，非网页版登录密码），但用腾讯邮箱的登录密码也能登录成功
my_pass = password.get()
# 收件人邮箱账号

my_user = receivior.get()
my_user=my_user.split(',',1)
def mail():
    try:
        errorscrem = tk.Text(frame, bg='white',  # 标签背景颜色
                             font=('微软雅黑', 9),  # 字体和字体大小
                             width=40, height=7,  # 标签长宽(以字符长度计算)
                             )
        frame.create_window(222,665,window=errorscrem)
        # 邮件内容
        mail_msg = text.get()
        content = MIMEText(mail_msg, 'html', 'utf-8')
        msg = MIMEMultipart()  # 多个MIME对象
        msg.attach(content)  # 添加内容
        msg['From'] = Header(send_name.get(), 'utf-8')  # 发件人
        msg['To'] = Header(rece_name.get(), 'utf-8')  # 收件人
        msg['Subject'] = Header(title.get(), 'utf-8')  # 主题
        # 附件
        #  attachment_1=MIMEText(open(r'C:/Users/zhengfei/Desktop/人.xlsx','rb').read(),'utf-8')
        # msg = MIMEMultipart()
        # attachment_1['content-Type']='application/octet-stream'
        # attachment_1['content-Disposition']='attachment;filename=r"C:/Users/zhengfei/Desktop/人.xlsx"'
        # msg.attach(attachment_1)
        if entry2.get()!="":
            file_name = entry2.get()  # 附件文件名
            file_path = os.path.join(file_name)  # 文件路径
            xlsx = MIMEApplication(open(file_path, 'rb').read())  # 打开Excel,读取Excel文件
            xlsx["Content-Type"] = 'application/octet-stream'  # 设置内容类型
            xlsx.add_header('Content-Disposition', 'attachment', filename=file_name)  # 添加到header信息

            msg.attach(xlsx)
        # SMTP服务器，腾讯企业邮箱端口是465，腾讯邮箱支持SSL(不强制)， 不支持TLS
        # qq邮箱smtp服务器地址:smtp.qq.com,端口号：456
        # 163邮箱smtp服务器地址：smtp.163.com，端口号：25
        server = smtplib.SMTP_SSL("smtp.exmail.qq.com", 465)
        # 登录服务器，括号中对应的是发件人邮箱账号、邮箱密码
        server.login(my_sender, my_pass)
        # 发送邮件，括号中对应的是发件人邮箱账号、收件人邮箱账号、发送邮件
        server.sendmail(my_sender, my_user, msg.as_string())
        # 关闭连接
        server.quit()
        errorscrem.insert(INSERT, '\n***************执行正常***************')
        # 如果 try 中的语句没有执行，则会执行下面的 ret=False
    except Exception as e:
        ret = False
        errorscrem.insert(INSERT, '\n***************程序报错，异常信息为:' + traceback.format_exc())
btn3=tk.Button(frame, text='执 行', command=mail,bg = "green",fg = "white")
frame.create_window(222,535,window=btn3)
window.mainloop()