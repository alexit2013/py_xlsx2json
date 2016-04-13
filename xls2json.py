__author__ = 'greatxin'
# -*- coding: utf-8 -*-

import json
import xlrd
import sys
import os

def start(xlsPath, outJsonPath=None):
    book = xlrd.open_workbook(xlsPath)
    sh = book.sheet_by_index(0)

    all = []
    for rowIndex in range(sh.nrows):
        if rowIndex <= 1:
            continue
        obj = {}
        hasKey = False
        for colIndex in range(sh.ncols):
            #check key
            key = sh.cell_value(rowx=1, colx=colIndex)
            if not key:
                continue
            obj[key] = sh.cell_value(rowx=rowIndex, colx=colIndex)
            hasKey = True
            # print key, obj[key]
        if hasKey:
            all.append(obj)

    # print all

    # output
    f = open(outJsonPath, "w")
    f.writelines("{")
    i = 0
    allLen = len(all)
    for obj in all:
        jsonStr = json.dumps(obj)
        f.writelines('"' + str(i) + '":' + jsonStr)
        if i < allLen - 1:
            f.writelines(",")
        i += 1
    f.writelines("}")
    f.close()

if __name__ == "__main__":
    args = sys.argv
    # start('data/test.xlsx', 'data/test.json')
    lenArgs = len(args)
    if lenArgs < 3:
        print('need two params, one excel , one output path.')
        sys.exit(-1)
    else:
        start(args[1], args[2])
        print ('success! json file saved to ' + args[2])
        sys.exit(0)