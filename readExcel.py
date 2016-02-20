import xlrd


def generate_words_to_num_dic():
    words_to_num_dic = {}
    i = 1
    words = [chr(i) for i in range(97, 123)]
    for word in words:
        words_to_num_dic[word.upper()] = i
        i += 1
    return words_to_num_dic


def trans_col_name_to_num(word, words_to_num_dic):
    # 把C列变成2,D列变成3
    number = words_to_num_dic[word.upper()]
    return number


def get_col_value_from_sheet(excel_file, target_col, index_num):
    # 打开excel,读取对应的列所有值,返回行数与值对应的list
    data = xlrd.open_workbook(excel_file)
    table = data.sheet_by_index(index_num - 1)
    print("Excel name:" + " ".join(data.sheet_names()) + ", we have " + str(data.nsheets) + " sheets," +
          "we are looking for No." + str(index_num) + " sheet")
    content_list = {}
    for row_num in range(table.nrows):
        content_list[row_num + 1] = table.cell_value(rowx=row_num, colx=target_col - 1).strip()
    return content_list


if __name__ == '__main__':
    excel = "分部分项工程量清单计价表.xls"
    col = "c"
    index = 1
    col_num = trans_col_name_to_num(col, generate_words_to_num_dic())
    print(get_col_value_from_sheet(excel, col_num, index))
