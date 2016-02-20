import xlrd


def get_col_value_from_sheet(sheet, target_col, index_num):
    # 打开excel,读取对应的列所有值,返回行数与值对应的list
    data = xlrd.open_workbook(sheet)
    table = data.sheet_by_index(index_num - 1)
    print("Excel name:" + " ".join(data.sheet_names()) + ", we have " + str(data.nsheets) + " sheets," +
          "we are looking for No." + str(index_num) + " sheet")
    content_list = {}
    for row_num in range(table.nrows):
        content_list[row_num + 1] = table.cell_value(rowx=row_num, colx=target_col - 1).strip()
    return content_list


print(get_col_value_from_sheet("分部分项工程量清单计价表.xls", 3, 1))
