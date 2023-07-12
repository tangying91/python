import openpyxl

# 打开 Excel 文件
from lib import my_xlsx

input_file_name = '../output/北区-2023年Q2销售业绩结算表-0703.xlsx'
wb = openpyxl.load_workbook(input_file_name)

# 遍历每个工作表
for sheet_name in wb.sheetnames:
    # 获取当前工作表
    sheet = wb[sheet_name]

    # 获取最大行数
    max_row = sheet.max_row

    # 遍历每一行，为第一列赋予自动编号
    column = my_xlsx.get_header_column_idx(sheet, '序号')
    if column != -1:
        # 因为第一行是标题，所以实际值行号往后加1
        for idx in range(1, max_row):
            sheet.cell(row=idx + 1, column=column).value = idx

# 保存文件
wb.save(input_file_name)
