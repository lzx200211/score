import smtplib
from email.mime.text import MIMEText
from email.header import Header

def send_email(sender_email, sender_password, receiver_email, subject, message):
    # 设置邮件内容和格式
    msg = MIMEText(message, 'plain', 'utf-8')
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = Header(subject, 'utf-8')

    try:
        # 连接QQ邮箱的SMTP服务器
        server = smtplib.SMTP_SSL('smtp.qq.com', 465)
        server.login(sender_email, sender_password)

        # 发送邮件
        server.sendmail(sender_email, [receiver_email], msg.as_string())
        print("邮件发送成功！")

    except Exception as e:
        print(f"发送邮件出错：{e}")

    finally:
        # 关闭连接
        server.quit()

# 测试发送邮件

