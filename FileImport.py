import tkinter as tk
from tkinter import filedialog
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.header import Header

def FileImport():
    # 创建一个Tkinter窗口
    root = tk.Tk()
    root.withdraw()

    # 弹出文件选择对话框
    file_path = filedialog.askopenfilename()

    # 打开工作簿
    workbook = openpyxl.load_workbook(file_path)
    print("import true!\n")

    # 获取工作表
    sheet = workbook.active

    analysis(sheet)
#文件内容保存在sheet

def analysis(sheet):
    name=None
    subject=None
    for i in range(3,999):
        num='C'+str(i)
        if name is None:
            name=sheet[num].value

        if sheet[num].value is None and name is not None:
            output(name,subject)
            break

        if sheet[num].value is not None:
            if name is not sheet[num].value :
                output(name,subject)
                name=sheet[num].value
                subject=None
                name=None

            num2="D"+str(i)
            num3="F"+str(i)
            if subject is None:
                subject=str(sheet[num2].value)+":"+str(sheet[num3].value)
            else:
                subject=str(subject)+str(sheet[num2].value)+":"+str(sheet[num3].value)

def output(name,subject):
    result="亲爱的"+str(name)+"同学:"+"\n"
    result=result+"祝贺您顺利完成本学期的学习！教务处在此向您发送最新的成绩单。\n"
    result=result+str(subject)+"\n"
    result=result+"希望您能够对自己的成绩感到满意，并继续保持努力和积极的学习态度。如果您在某些科目上没有达到预期的成绩，不要灰心，这也是学习过程中的一部分。我们鼓励您与您的任课教师或辅导员进行交流，他们将很乐意为您解答任何疑问并提供帮助。请记住，学习是一个持续不断的过程，我们相信您有能力克服困难并取得更大的进步。"+"\n"
    result=result+"再次恭喜您，祝您学习进步、事业成功！\n"
    result=result+"教务处\n"
    send_email('2505565012@qq.com', '20021003cnwby', '2194958319@qq.com', "成绩通知", result)
    print(result)#邮件内容在result

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
FileImport()