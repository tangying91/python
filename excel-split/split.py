from lib import my_xlsx

# 拆分表格的关键字
split_keys = ['苏皖大区', '上海分公司', '北京分公司', '北区', '西区', '物流交通', '南区', '浙江大区', '福建办事处']
# 最终需要保留的数据分页
keep_sheets = ['大区', '办事处', '组长', '个人', 'Q2业绩清单', '剔除清单']
# 需要处理的源文件
input_file_name = './input/2023年Q2业绩综合分析-7.3.xlsx'
# 拆分后文件输出目录
output_file_path = './output/'
# 拆分后的文件名称后缀
output_file_suffix = '-2023年Q2销售业绩结算表-0703'

# 开始处理Excel表格
my_xlsx.excel_split(input_file_name, split_keys, keep_sheets, output_file_path, output_file_suffix)

