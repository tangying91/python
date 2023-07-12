import xlwings as xw

# 尝试使用xlwings

# 打开Excel文件
wb = xw.Book("your_file.xlsx")
sheet = wb.sheets[0]

# 获取数据范围
data_range = sheet.range("A1").expand()

# 将公式替换为值
for cell in data_range:
    if cell.formula:
        cell.value = cell.value

# 保存修改后的Excel文件
wb.save("output_file.xlsx")
wb.close()