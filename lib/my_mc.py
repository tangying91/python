
# 根据销售部门获取大区
def get_area_by_department(book_sheet, department):
    for row in book_sheet.iter_rows(min_row=2):
        for cell in row:
            if department == cell.value:
                return book_sheet.cell(row=1, column=cell.column).value
    return ""


# 根据销售部门获取大区
def get_all_areas(book_sheet):
    areas = []
    for row in book_sheet.iter_rows(min_row=1, max_row=1):
        for cell in row:
            areas.append(cell.value)
    return areas
