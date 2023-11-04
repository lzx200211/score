import pandas as pd

def import_excel(file_path):
    try:
        df = pd.read_excel("./成绩表.xlsx")  # 使用pandas库的read_excel函数读取Excel文件
        return df
    except Exception as e:
        print("导入Excel文件失败：", str(e))
        return None

# 示例用法
file_path = "./成绩表.xlsx"  # Excel文件路径
data = import_excel(file_path)  # 调用导入函数
if data is not None:
    print("成功导入Excel数据：")
    print(data)


