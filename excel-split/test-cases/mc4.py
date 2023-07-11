import pandas as pd

# 打开 Excel 文件
excel_file = pd.ExcelFile('北京分公司.xlsx')

# 创建新的 Excel 文件
output_file = pd.ExcelWriter('新文件名.xlsx')

# 遍历每一页
for sheet_name in excel_file.sheet_names:
    # 读取当前页的数据
    df = pd.read_excel(excel_file, sheet_name=sheet_name)

    # 删除已有的序号列
    if '序号' in df.columns:
        df.drop('序号', axis=1, inplace=True)

    # 添加新的序号列
    df.insert(0, '序号', range(1, len(df) + 1))

    # 写入当前页的数据到新的 Excel 文件中
    df.to_excel(output_file, sheet_name=sheet_name, index=False)

# 保存并关闭新的 Excel 文件
output_file.save()
output_file.close()

# 关闭原始 Excel 文件
excel_file.close()
