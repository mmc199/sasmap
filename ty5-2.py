import pandas as pd
import os
import sys
from collections import defaultdict

if getattr(sys, 'frozen', False):
    # 打包程序
    exe_dir = os.path.dirname(sys.executable)
else:
    # 非打包程序
    exe_dir = os.path.dirname(os.path.abspath(__file__))

# 获取当前脚本所在的目录
script_dir = exe_dir
# 将工作目录设置为当前脚本所在的目录
os.chdir(script_dir)

# 读取Excel文件
excel_path = os.path.join(script_dir, 'ty5-2.xls')
df = pd.read_excel(excel_path, sheet_name=[0, 1])

# 处理第一个表格（参数信息）
df1 = df[0].iloc[2:, :]
data_dict = defaultdict(list)

for i in range(len(df1)):
    for j in range(len(df1.columns)):
        if pd.notnull(df1.iloc[i,j]) and '=' in str(df[0].iloc[1,j]) and 'N' not in str(df1.iloc[i,1]):
            key = str(df1.iloc[i, 0]).strip()
            value = str(df1.iloc[i,j])
            data_dict[key].append(value)

# 处理第二个表格（URL脚本信息）
df2 = df[1].iloc[2:, :]
geturl_dict = defaultdict(list)

for i in range(len(df2)):
    for j in range(len(df2.columns)):
        if pd.notnull(df2.iloc[i,j]) and '=' not in str(df[1].iloc[1,j]) and 'N' not in str(df2.iloc[i,1]):
            key = str(df2.iloc[i, 0]).strip()
            value = str(df2.iloc[i,j])
            geturl_dict[key].append(value)

# 遍历data_dict和geturl_dict，创建文件和目录
for dict_data, file_name in [(data_dict, 'params.txt'), (geturl_dict, 'GetUrlScript.txt')]:
    for key, content_list in dict_data.items():
        directory_name = f"{key}.zmp"
        directory_path = os.path.join(script_dir, directory_name)
        
        if not os.path.exists(directory_path):
            os.makedirs(directory_path)
        
        file_path = os.path.join(directory_path, file_name)
        
        with open(file_path, 'w', encoding='gbk') as f:
            f.write('\n'.join(content_list))
            f.write('\n\n')

print("使用的编码: gbk")
print("处理完成。")