#!/usr/bin/python3
# -*- coding: UTF-8 -*-

import pandas as pd

# 读取 Excel 文件中的所有工作表
sheets_dict = pd.read_excel('2023年Q1业绩综合分析V4终版-4.11.xlsx', sheet_name=None)

# 定义字符串数组
strings = ['苏皖大区', '上海分公司', '北京分公司', '北区', '西区', '物流交通大区', '南区', '浙江大区']

# 遍历每个字符串
for string in strings:
    # 创建一个空的字典来保存筛选后的工作表数据
    filtered_sheets_dict = {}

    # 遍历每个工作表
    for sheet_name, df in sheets_dict.items():
        # 筛选含有指定字符串的行数据
        filtered_df = df[df.astype(str).apply(lambda row: any(string in cell for cell in row), axis=1)]
        # 将筛选后的数据添加到结果字典中
        filtered_sheets_dict[sheet_name] = filtered_df

    # 导出筛选后的数据到 Excel 文件，每个工作表分别保存为一个 sheet
    with pd.ExcelWriter(f'{string}.xlsx') as writer:
        for sheet_name, filtered_df in filtered_sheets_dict.items():
            filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)


