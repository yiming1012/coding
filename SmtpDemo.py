# -*- coding: utf-8 -*-
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# 第三方 SMTP 服务
mail_host = "smtp.qq.com"  # 设置服务器
mail_user = "XXXX@qq.com"  # 用户名
mail_pass = "XXXX"  # 口令,QQ里面的授权码
receivers = ['XXXX@qq.com','XXXX@qq.com']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱


msg = MIMEMultipart() # 创建一个带附件的实例
msg["Subject"] = "Tomorrow is another day"
msg["From"] = mail_user
msg["To"] = ','.join(receivers)

# ---文字部分---
part = MIMEText("请查收，谢谢！")
msg.attach(part)

# ---附件部分---
part = MIMEApplication(open('foo.xlsx', 'rb').read())
part.add_header('Content-Disposition', 'attachment', filename="foo.xlsx")
msg.attach(part)

try:
    s = smtplib.SMTP("smtp.qq.com", timeout=30)  # 连接smtp邮件服务器,端口默认是25
    s.login(mail_user, mail_pass)  # 登陆服务器
    s.sendmail(mail_user, receivers, msg.as_string())  # 发送邮件
    s.close()
except Exception as e:
    print("error:", e)
