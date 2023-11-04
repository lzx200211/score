import pandas as pd
import numpy as np

def import_excel(file_path):
    try:
        df = pd.read_excel(file_path)  # 使用pandas库的read_excel函数读取Excel文件
        df = df.replace('\xa0','',regex=True)
        data = df
        data = np.array(data)
        student = []
        k = 0
        for i in range(0, len(data)):
            name = data[i][2]
            course = data[i][3]
            grade = data[i][6]
            if name == data[i - 1][2]:
                student[k - 1][1].append(course + ':' + str(grade))
            else:
                student.append([name, [course + ':' + str(grade)]])
                k += 1
        return student
    except Exception as e:
        print("导入Excel文件失败：", str(e))
        return None

# 示例用法
file_path = "D:\PythonPJ\Software Engineering\Text4_1\成绩表.xlsx"  # Excel文件路径
data = import_excel(file_path)  # 调用导入函数



