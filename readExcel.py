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
    content_list = {}
    for row_num in range(table.nrows):
        content_list[row_num + 1] = table.cell_value(rowx=row_num, colx=target_col - 1).strip()
    return content_list


def find_diff(string1, string2):
    col_content = ""
    if len(string1) > len(string2) or len(string1) == len(string2):
        for word in string1:
            if word not in string2:
                col_content += "\033[1;31;40m%s\033[0m" % str(word)
            else:
                col_content += word
    else:
        for word in string2:
            if word not in string1:
                col_content += "\033[1;31;40m%s\033[0m" % str(word)
            else:
                col_content += word
    return col_content


def diff_two_excel(excel_1, excel_2, col_num, index):
    diff_content = {}
    excel_1_content = get_col_value_from_sheet(excel_1, col_num, index)
    excel_2_content = get_col_value_from_sheet(excel_2, col_num, index)
    for key in excel_1_content:
        if excel_1_content[key] != excel_2_content[key]:
            diff_content[key] = find_diff(excel_1_content[key], excel_2_content[key])
    return diff_content


def nice_print(dic):
    for key in dic:
        print("Line: " + str(key) + " diff: " + dic[key])

if __name__ == '__main__':
    excel_1 = "分部分项工程量清单计价表2.xlsx"
    excel_2 = "分部分项工程量清单计价表.xls"
    col = "c"
    index = 1
    col_num = trans_col_name_to_num(col, generate_words_to_num_dic())
    print("Different contents in <" + excel_1 + "> and <" + excel_2 + ">")
    nice_print(diff_two_excel(excel_1, excel_2, col_num, index))
