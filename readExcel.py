import xlrd


# book = xlrd.open_workbook("分部分项工程量清单计价表.xls")
# print ("The number of worksheets is" , book.nsheets)
# print ("Worksheet name(s):" , book.sheet_names())
# sh = book.sheet_by_index(0)
# print (sh.name, sh.nrows, sh.ncols)
# print ("Cell D30 is", sh.cell_value(rowx=29, colx=2))
# for rx in range(sh.nrows):
#  print ("row: "+ str(rx + 1) + " " + "value: " + sh.cell_value(rowx=rx, colx=2))


def get_col_value_from_sheet(sheet, target_col, index_num):
    data = xlrd.open_workbook(sheet)
    table = data.sheet_by_index(index_num - 1)
    print("Excel name:" + " ".join(data.sheet_names()) + ", we have " + str(data.nsheets) + " sheets," +
          "we are looking for No." + str(index_num) + " sheet")
    content_list = {}
    for row_num in range(table.nrows):
        content_list[row_num + 1] = table.cell_value(rowx=row_num, colx=target_col - 1).strip()
    return content_list


print(get_col_value_from_sheet("分部分项工程量清单计价表.xls", 3, 1))
