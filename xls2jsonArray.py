__author__ = 'greatxin'
# -*- coding: utf-8 -*-

import json
import xlrd
import sys

def start(xlsPath, outJsonPath=None):
    book = xlrd.open_workbook(xlsPath)
    # print "The number of worksheets is", book.nsheets
    # print "Worksheet name(s):", book.sheet_names()

    sh = book.sheet_by_index(0)

    all = []
    for rowIndex in xrange(sh.nrows):
        all.append(sh.row_values(rowIndex))
    # print all
    # output
    f = open(outJsonPath, "w")
    str = json.dumps(all)
    f.write(str)
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