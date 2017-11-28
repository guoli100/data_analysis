#!/usr/bin/env python3
import os
import sys
import glob
import pandas as pd


def pdm_dmp_anlys(xl):
    if not os.path.isfile(xl):
        sys.exit(xl + ' is not a file')

    filename, ext = os.path.splitext(xl)
    df = pd.read_excel(xl, usecols="C, E")
    print(df)


if __name__ == '__main__':
    if len(sys.argv) >= 2:
        filelist = []
        for f in sys.argv[1:]:
            filelist.extend(glob.glob(f))
        print(filelist)

        for f in filelist:
            # 将文件名传入你的处理函数
            print('\n正在分析文件: ' + f)
            pdm_dmp_anlys(f)
    else:
        sys.exit('需要指定一个或多个Excel文件')
