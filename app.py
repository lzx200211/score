from flask import Flask, request, jsonify, render_template
from werkzeug.utils import secure_filename
import os
import smtplib
from email.mime.text import MIMEText
from email.header import Header
import pandas as pd
import numpy as np

# 获取当前文件的目录
app_folder = os.path.dirname(__file__)
print("Current directory:", app_folder)

app = Flask(__name__)

# 配置文件上传的文件夹
app.config['UPLOAD_FOLDER'] = os.path.join(app_folder, 'uploads')
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

# 辅助函数：检查文件扩展名是否允许
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

# 导入Excel文件
def import_excel(file_path):
    if not os.path.exists(file_path):
        print(f"文件 '{file_path}' 不存在")
        return None
    try:
        df = pd.read_excel(file_path)  # 使用pandas库的read_excel函数读取Excel文件
        df = df.replace('\xa0', '', regex=True)
        data = df
        data = np.array(data)

        student = []

        k = 0
        for i in range(1, len(data)):
            name = data[i][2]
            course = data[i][3]
            grade = data[i][6]
            if name == data[i - 1][2]:
                student[k - 1][1].append(course + ':' + str(grade))
            else:
                student.append([name, [course + str(grade)]])
                k += 1
        return student
    except Exception as e:
        print("导入Excel文件失败：", str(e))
        return None


@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'})

    file = request.files['file']

    if file.filename == '':
        print(file.filename)
        return jsonify({'error': 'No selected file'})

    if file and allowed_file(file.filename):
        filename = file.filename  # 使用原始文件名
        #file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        # 设置上传文件的保存路径为当前目录下的 'uploads' 文件夹
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        return jsonify({'message': 'File uploaded successfully'})
    else:
        return jsonify({'error': 'Invalid file type'})

    
    receiver_email = '3475869824@qq.com'

@app.route('/send_email', methods=['POST'])
def send_email(sender_email='2449165916@qq.com', sender_password = 'gfgkprtzfibuebbj', receiver_email = '3475869824@qq.com', subject = '成绩报告单' ):
    stu = import_excel(os.path.join(app.config['UPLOAD_FOLDER'], '成绩表.xlsx'))
    # print(stu)
    # print(len(stu))
    k =len(stu)
    sender_email = '2449165916@qq.com'
    sender_password = 'gfgkprtzfibuebbj'
    receiver_email = '1161477176@qq.com'
    for i in range(k):
        # print(k)
        # print(i)
        name = stu[i][0]
        subject = '成绩报告单'
        results = stu[i][1]
        message = f'''亲爱的{name}同学:

        祝贺您顺利完成本学期的学习！教务处在此向您发送最新的成绩单。

        {results}

        ……

        希望您能够对自己的成绩感到满意，并继续保持努力和积极的学习态度。如果您在某些科目上没有达到预期的成绩，不要灰心，这也是学习过程中的一部分。我们鼓励您与您的任课教师或辅导员进行交流，他们将很乐意为您解答任何疑问并提供帮助。请记住，学习是一个持续不断的过程，我们相信您有能力克服困难并取得更大的进步。

        再次恭喜您，祝您学习进步、事业成功！

        教务处'''

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
            server.quit()
            print('邮件发送成功！')

        except Exception as e:
            return jsonify({'error': f"发送邮件出错：{e}"})

    return jsonify({'message': '所有邮件已成功发送！'})
    



if __name__ == '__main__':
    app.run(debug=True)
