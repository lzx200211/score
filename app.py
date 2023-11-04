from flask import Flask, request, jsonify, render_template
from werkzeug.utils import secure_filename
import os
import smtplib
from email.mime.text import MIMEText
from email.header import Header
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

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'})

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': 'No selected file'})

    if file and allowed_file(file.filename):
        filename = file.filename  # 使用原始文件名
        #file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        # 设置上传文件的保存路径为当前目录下的 'uploads' 文件夹
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        return jsonify({'message': 'File uploaded successfully'})
    else:
        return jsonify({'error': 'Invalid file type'})

@app.route('/send_email', methods=['POST'])
def send_email():
    # sender_email = 'your_sender_email@qq.com'  # 发件人邮箱
    # sender_password = 'your_sender_password'  # 发件人邮箱密码
    # receiver_email = 'receiver_email@example.com'  # 收件人邮箱
    # subject = '邮件主题'
    # message = '邮件内容'



    sender_email = '2449165916@qq.com'
    sender_password = 'gfgkprtzfibuebbj'
    receiver_email = '3475869824@qq.com'
    subject = '邮件主题'
    message = '邮件内容'

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

        return jsonify({'message': '邮件发送成功！'})

    except Exception as e:
        return jsonify({'error': f"发送邮件出错：{e}"})



if __name__ == '__main__':
    app.run(debug=True)
