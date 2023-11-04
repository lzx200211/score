from flask import Flask, request, jsonify, render_template
from werkzeug.utils import secure_filename
import os
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

if __name__ == '__main__':
    app.run(debug=True)
