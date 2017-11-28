#!/usr/bin/env python3
'''
对处理过的oracle dump文件进行表结构提取，并将结果写入Excel
'''
import re
import os
import sys
import glob
import pandas as pd


def parse_dmp(file):
    if not os.path.isfile(file):
        sys.exit(file + ' is not a file')

    filename, ext = os.path.splitext(file)

    pd.set_option('display.height', 1000)
    pd.set_option('display.max_rows', 500)
    pd.set_option('display.max_columns', 500)
    pd.set_option('display.width', 1000)

    with open(file, 'r', encoding='utf-8') as dmpfile:
        data = []
        for line in dmpfile:
            tab_name = re.match(r'CREATE TABLE "(.*)" \(', line).group(1)
            m_col = re.search(r'\((.*)\)', line)
            cols = m_col.group(1).split('|')

            for col in cols:
                ncol = col.strip().split(' ', 2)
                ncol[0] = ncol[0].strip('"')
                ncol.insert(0, tab_name)
                data.append(ncol)

        df = pd.DataFrame(data, columns=['表名', '字段名', '字段类型', '非空'])
        df.to_excel(filename + '.xlsx')


if __name__ == '__main__':
    if len(sys.argv) >= 2:
        filelist = []
        for f in sys.argv[1:]:
            filelist.extend(glob.glob(f))
        print(filelist)

        for f in filelist:
            # 将文件名传入你的处理函数
            parse_dmp(f)
    else:
        sys.exit('需要指定一个或多个格式化后的dmp文件')
