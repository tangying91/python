import openpyxl

# 删除空分页
# 判断第二行是否为空或没有数据

# 打开 Excel 文件
workbook = openpyxl.load_workbook('./files/北京分公司-2023年Q2销售业绩结算表-0703.xlsx')

# 获取所有的工作表
worksheets = workbook.sheetnames

# 遍历每个工作表
for sheet_name in worksheets:
    sheet = workbook[sheet_name]

    # 判断第二行是否为空或没有数据
    row_values = [cell.value for cell in sheet[2]]
    if all(value is None for value in row_values):
        # 删除该工作表
        workbook.remove(sheet)

# 保存修改后的文件
workbook.save('北京分公司.xlsx')
