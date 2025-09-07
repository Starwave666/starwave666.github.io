import pandas as pd
import os

# 读取Excel文件
file_path = 'MCK 客--2024年 & 2025年 GI & GF 订单销量对比数据 更新.xlsx'
sheet_name = 'Sheet1'  # 工作表1

# 读取数据
df = pd.read_excel(file_path, sheet_name=sheet_name)

# 计算每份表格的行数
total_rows = len(df)
rows_per_split = total_rows // 4

# 创建保存分拆文件的目录
output_dir = 'split_data'
os.makedirs(output_dir, exist_ok=True)

# 分拆数据并保存为4个文件
for i in range(4):
    start_row = i * rows_per_split
    if i == 3:  # 最后一份包含所有剩余数据
        end_row = total_rows
    else:
        end_row = (i + 1) * rows_per_split
    
    split_df = df.iloc[start_row:end_row]
    split_df.to_excel(os.path.join(output_dir, f'split_sheet_{i+1}.xlsx'), index=False)

print("Excel文件已成功分拆为4份表格，保存在split_data目录中")