import os  
import zipfile  
import pandas as pd  
import shutil  
  
# 创建csv和xlsx文件夹（如果不存在）  
csv_dir = 'csv'  
xlsx_dir = 'xlsx'  
os.makedirs(csv_dir, exist_ok=True)  
os.makedirs(xlsx_dir, exist_ok=True)  
  
# 查找并解压所有的zip文件到csv文件夹  
for filename in os.listdir('.'):  
    if filename.endswith('.zip'):  
        zip_path = os.path.join('.', filename)  
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:  
            # 提取zip文件到csv文件夹，并保留原文件名（不含.zip扩展名）  
            csv_file_without_ext = os.path.splitext(filename)[0]  
            zip_ref.extractall(os.path.join(csv_dir, csv_file_without_ext))  
  
# 扫描csv文件夹下的csv文件，转换为xlsx并保存到xlsx文件夹  
for root, dirs, files in os.walk(csv_dir):  
    for file in files:  
        if file.endswith('.csv'):  
            # 构建csv文件的完整路径  
            csv_path = os.path.join(root, file)  
              
            # 提取文件名中的数字部分作为xlsx文件名  
            xlsx_filename = os.path.splitext(file)[0].split('_')[0] + '.xlsx'  
              
            # 构建xlsx文件的完整路径  
            xlsx_path = os.path.join(xlsx_dir, xlsx_filename)  
              
            # 读取csv文件并保存到xlsx文件  
            df = pd.read_csv(csv_path)  
            df.to_excel(xlsx_path, index=False, engine='openpyxl')  
              
            # 可选：删除已转换的csv文件（如果需要）  
            # os.remove(csv_path)  
  
# 可选：删除csv文件夹（如果所有csv文件都已转换且不再需要）  
shutil.rmtree(csv_dir)
  
# 可选：删除所有已解压的zip文件（如果需要）  
for filename in os.listdir('.'):  
    if filename.endswith('.zip'):  
        os.remove(filename)