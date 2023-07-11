#!/usr/bin/python3
# -*- coding: UTF-8 -*-

import openpyxl
import pandas as pd
from lib import my_xlsx as xutils

# log
print("程序开始运行...")

# 解析excel
wb = openpyxl.load_workbook("销管中心-合同管理表20210713-3.xlsx", data_only=True)
book_sheet = wb.get_sheet_by_name("美创集团")

# log
print("目标文件解析成功...")

headers = ["项目编号", "合同编号", "销售团队", "销售", "合同金额"]
datas = xutils.filter_sheet_data(book_sheet, headers, {'销售': '杨楠'})

# 封装结果
result_datas = {}
for data in datas:
    key = data["合同编号"]
    # 数据清洗
    if "2021" not in key or "-2020" in key:
        continue

    # 正常数据
    if key not in result_datas.keys():
        result_datas[key] = data
    else:
        try:
            d = result_datas[key]
            p = d["合同金额"] + data["合同金额"]
            r = {"合同金额": p}
            d.update(r)
        except TypeError:
            print(data)
        except ValueError:
            print(data)
        else:
            continue

# 数据包装
df = pd.DataFrame(result_datas)

# 保存 dataframe
df.to_csv('2021合同清单.csv')
