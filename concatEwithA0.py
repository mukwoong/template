import os
import pandas as pd

# 指定文件夹路径
folder_path = r'C:\Users\Thinkpad\Desktop\1'

# 初始化一个空的列表来存储每个读取的DataFrame
dataframes = []

# 遍历文件夹中的所有文件
for filename in os.listdir(folder_path):
    # 检查文件是否为Excel文件（可以根据实际情况扩展）
    if filename.endswith('.xlsx') or filename.endswith('.xls'):
        file_path = os.path.join(folder_path, filename)
        # 读取Excel文件中的第一个表
        df = pd.read_excel(file_path)
        # 将读取的DataFrame添加到列表中
        dataframes.append(df)

# 使用concat函数连接所有的DataFrame
result = pd.concat(dataframes)

# 输出到新的Excel文件
output_path = os.path.join(folder_path, 'output.xlsx')
result.to_excel(output_path, index=False)

print(f"All files have been concatenated and saved to {output_path}")
