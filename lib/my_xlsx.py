import warnings

# 过滤无意义的警告
warnings.filterwarnings('ignore')


# 根据header找到列的下标，从 0 开始
def get_header_column_idx(book_sheet, header):
    for row in book_sheet.iter_rows(min_row=1, max_row=1):
        for cell in row:
            if header == cell.value:
                return cell.column - 1
    return ""


# 获取某一列的不重复数据
def get_no_repeat_column_data(book_sheet, title):
    column = get_header_column_idx(book_sheet, title)
    datas = []
    for row in book_sheet.iter_rows(min_row=2):
        datas.append(row[column].value)
    return datas


# 获取指定分页，指定列，并过滤关键字的数据
# headers是数组，filters是字典
def filter_sheet_data(book_sheet, headers, filters):
    title_columns = {}
    for header in headers:
        column = get_header_column_idx(book_sheet, header)
        title_columns[column] = header

    datas = []
    for row in book_sheet.iter_rows(min_row=2):
        data = {}
        append_data = True
        all_none = True
        for column, header in title_columns.items():
            value = row[column].value
            data[header] = value

            # 检查是否需要过滤
            if header in filters.keys() and value == filters[header]:
                append_data = False

            # 检查是否整行都是空数据
            if value is not None:
                all_none = False

        if append_data is True and all_none is False:
            datas.append(data)
    return datas


# 设置分页指定列列宽
def set_sheet_column_width(book_sheet, columns, width):
    for column in columns:
        book_sheet.column_dimensions[column].width = width
