#!/usr/bin/env python3
import sys
import pandas as pd

pd.set_option('display.height', 1000)
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)


def merge_pdm_xls(filelist):
    dflist = []
    for f in filelist:
        #  dflist.append(pd.read_excel(f, usecols='C, E'))
        dflist.append(pd.read_excel(f))
    return pd.concat(dflist)


def merge_dmp_xlsx(filelist):
    dflist = []
    for f in filelist:
        dflist.append(pd.read_excel(f, usecols='B, C'))
    return pd.concat(dflist)


if __name__ == '__main__':
    if len(sys.argv) >= 2:
        # 读取所有pdm的excel文件到DataFrame
        pdm_data = merge_pdm_xls(sys.argv[1:-2])
        # 读取dmp的excel文件到DataFrame
        #  dmp_data = merge_dmp_xlsx(sys.argv[-2:])
        dmp_data = pd.read_excel(sys.argv[-1], usecols='B, C')

        # pdm_data和dmp_data连接用的键
        keys = ['表', '字段']

        # 重命名pdm_data的列标签名，以方便连接
        cols = list(pdm_data.columns)
        cols[1] = '表'
        cols[3] = '字段'
        pdm_data.columns = pd.Index(cols)

        # 重命名dmp_data的列标签名，以方便连接
        dmp_data.columns = pd.Index(keys)

        # 以连接键对pdm_data和dmp_data去重
        pdm_data = pdm_data.drop_duplicates(keys)
        dmp_data = dmp_data.drop_duplicates(keys)
        pd.merge(pdm_data, dmp_data).to_excel('pdm_dmp_匹配.xlsx')
    else:
        sys.exit('需要指定一个或多个Excel文件')
