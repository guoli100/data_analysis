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
        dflist.append(pd.read_excel(f, usecols='C, E'))
    return pd.concat(dflist)


def merge_dmp_xlsx(filelist):
    dflist = []
    for f in filelist:
        dflist.append(pd.read_excel(f, usecols='B, C'))
    return pd.concat(dflist)


if __name__ == '__main__':
    if len(sys.argv) >= 2:
        pdm_data = merge_pdm_xls(sys.argv[1:-2])
        dmp_data = merge_dmp_xlsx(sys.argv[-2:])
        pdm_data.columns = pd.Index(['表', '字段'])
        dmp_data.columns = pd.Index(['表', '字段'])
        pdm_data = pdm_data.drop_duplicates()
        dmp_data = dmp_data.drop_duplicates()
        print(pd.merge(pdm_data, dmp_data))
        #  print(dmp_data)
    else:
        sys.exit('需要指定一个或多个Excel文件')
