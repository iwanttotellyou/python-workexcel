# -*- coding: utf-8 -*-
import xlrd
import json


def open_excel(file_name):
    '''
    打开excel
    :param file_name: excel的名称
    :return: xlrd.book.Book对象
    '''
    try:
        data = xlrd.open_workbook(file_name)
        return data
    except Exception, e:
        print str(e)


def excel_table_byindex(file_name=None, col_names_index=0, by_index=0):
    '''
    根据表的索引获取数据
    :param file_name: Excel文件路径
    :param col_names_index: 表头列名所在行的索引
    :param by_index: 表的索引
    :return: 每行的字典list
    '''
    if file_name is None:
        file_name = 'file.xlsx'
    data = open_excel(file_name)
    table = data.sheets()[by_index]
    nrows = table.nrows  # 行数
    colnames = table.row_values(col_names_index)  # 获取字典的key
    list = []
    for rownum in xrange(1, nrows):
        row = table.row_values(rownum)
        if row:
            app = {}
            for i in xrange(len(colnames)):
                app[colnames[i]] = row[i]
            list.append(app)
    return list


def main():
    tables = excel_table_byindex()
    json_list = json.dumps(tables)
    print json_list


if __name__ == "__main__":
    main()