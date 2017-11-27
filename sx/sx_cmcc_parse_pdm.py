#!/usr/bin/env python3

import sys
import os
import glob
import xml.etree.ElementTree as ET
import pandas as pd


def contain_chinese(strings):
    for s in strings:
        if '\u4e00' <= s <= '\u9fff':
            return True
    return False


def xml2excel(xml):
    '''
    读取pdm文件，从中抽取所有表信息及字段信息，将信息写入两个Excel文件中。
    '''
    if not os.path.isfile(xml):
        sys.exit(xml + ' is not a file')

    filename, ext = os.path.splitext(xml)

    table_data = pd.DataFrame(columns=[
        '表中文名',
        '表英文名',
        '中文缺失',
        '乱码',
        '表中文描述'])
    col_data = pd.DataFrame(columns=[
        '表中文名',
        '表英文名',
        '字段中文名',
        '字段英文名',
        '字段类型',
        '主键',
        '不能为空',
        '中文缺失',
        '乱码',
        '字段描述'])

    ns = {'a': 'attribute',
          'c': 'collection',
          'o': 'object'}

    ET.register_namespace('a', 'attribute')
    ET.register_namespace('c', 'collection')
    ET.register_namespace('o', 'object')

    tree = ET.parse(xml)

    j = 0
    for i, table in enumerate(tree.findall('.//c:Tables/o:Table', ns)):
        # 抽取表信息
        t_name = table.find('./a:Name', ns)
        t_code = table.find('./a:Code', ns)
        t_comment = table.find('./a:Comment', ns)
        table_data.loc[i, '表中文名'] = t_name.text
        table_data.loc[i, '表英文名'] = t_code.text
        if not contain_chinese(t_name.text):
            table_data.loc[i, '中文缺失'] = t_name.text
        if '�' in t_name.text:
            table_data.loc[i, '乱码'] = t_name.text
        if t_comment is not None:
            table_data.loc[i, '表中文描述'] = t_comment.text

        # 抽取主键信息
        pk = table.find('./c:PrimaryKey/o:Key', ns)
        pk_col_ids = []
        if pk is not None:
            for pk_column in table.findall('./c:Keys/'
                                           'o:Key[@Id=\'' + pk.attrib['Ref'] + '\']//'
                                           'o:Column', ns):
                pk_col_ids.append(pk_column.attrib['Ref'])

        # 抽取字段信息
        for col in table.findall('.//c:Columns/o:Column', ns):
            c_name = col.find('./a:Name', ns)
            c_code = col.find('./a:Code', ns)
            c_dtype = col.find('./a:DataType', ns)
            c_mand = col.find('./a:Mandatory', ns)
            c_comment = col.find('./a:Comment', ns)

            col_data.loc[j, '表中文名'] = t_name.text
            col_data.loc[j, '表英文名'] = t_code.text
            col_data.loc[j, '字段中文名'] = c_name.text
            col_data.loc[j, '字段英文名'] = c_code.text
            col_data.loc[j, '字段类型'] = c_dtype.text
            if col.attrib['Id'] in pk_col_ids:
                col_data.loc[j, '主键'] = 'Yes'
            if c_mand is not None and c_mand.text == '1':
                col_data.loc[j, '不能为空'] = 'True'
            if not contain_chinese(c_name.text):
                col_data.loc[j, '中文缺失'] = c_name.text
            if '�' in c_name.text:
                col_data.loc[j, '乱码'] = c_name.text
            if c_comment is not None:
                col_data.loc[j, '字段描述'] = c_comment.text
            j += 1

    table_data.to_excel(filename + '_表.xls')
    col_data.to_excel(filename + '_字段.xls')

    print('抽取成功，数据已写入到：')
    print('  ' + filename + '_表.xls ')
    print('  ' + filename + '_字段.xls')


if __name__ == '__main__':
    if len(sys.argv) >= 2:
        filelist = []
        for f in sys.argv[1:]:
            filelist.extend(glob.glob(f))
        print(filelist)

        for f in filelist:
            print('\n正在抽取文件: ' + f)
            xml2excel(f)
    else:
        sys.exit('需要指定一个或多个XML文件')
