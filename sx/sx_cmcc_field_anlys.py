#!/usr/bin/env python3

import os
import sys
import glob
import pandas as pd


def name_merge(s):
    return '，'.join(list(set(s)))


def field_anlys(xl):
    if not os.path.isfile(xl):
        sys.exit(xl + ' is not a file')

    filename, ext = os.path.splitext(xl)
    df = pd.read_excel(xl)

    # 英文名和类型统一大写
    df['字段英文名'] = df['字段英文名'].str.upper()
    df['字段类型'] = df['字段类型'].str.upper()

    data_en = df.groupby('字段英文名')[['字段中文名', '字段类型']].agg(name_merge)
    data_en['多个中文名'] = data_en['字段中文名'].str.contains('，')
    data_en['多个类型'] = data_en['字段类型'].str.contains('，')
    data_en = data_en.reindex(columns=['字段中文名', '多个中文名', '字段类型', '多个类型'])
    data_cn = df.groupby('字段中文名')['字段英文名'].agg(name_merge)
    data_cn = pd.DataFrame(data_cn)
    data_cn['多个英文名'] = data_cn['字段英文名'].str.contains('，')

    writer = pd.ExcelWriter(filename + '分析.xls')
    data_en.to_excel(writer, '字段英文名')
    data_cn.to_excel(writer, '字段中文名')
    writer.save()


if __name__ == '__main__':
    if len(sys.argv) >= 2:
        filelist = []
        for f in sys.argv[1:]:
            filelist.extend(glob.glob(f))
        print(filelist)

        for f in filelist:
            # 将文件名传入你的处理函数
            print('\n正在分析文件: ' + f)
            field_anlys(f)
    else:
        sys.exit('需要指定一个或多个Excel文件')
