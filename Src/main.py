#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlrd
import xlwt
import re

excel_path = '../target.xls'
output_path = '../output.xls'
CP_Index = 17

# 向表table的第row行，写入数组valList的数据
def Excel_WriteLine(table,row,val):
    if isinstance(val,str):
        table.write(row, 0, val)
    else:
        for i in range(len(val)):
            table.write(row, i, val[i])

    row += 1
    return row

def Excel_Process():
    # 获得excel 第一个sheet的内容
    in_file = xlrd.open_workbook(excel_path)
    in_table = in_file.sheets()[0]

    # 获得行列数
    in_rows = in_table.nrows
    in_cols = in_table.ncols

    # 建立用于写的excel
    out_file = xlwt.Workbook()
    out_table = out_file.add_sheet('sheet1',cell_overwrite_ok=True)

    # 写入首行的索引
    row_index = 0
    row_index = Excel_WriteLine(out_table,row_index,in_table.row_values(0))

    pattern = re.compile(r'.+?-.{1,2}\s')

    # 逐行读数据
    for i in range(1,in_rows):
        lineArray = in_table.row_values(i)
        # CP列存在有效数据
        if lineArray[CP_Index].strip() != '':
            TmpArray =  lineArray.copy()
            cpVal = lineArray[CP_Index].strip(' ;')
            # 以;拆分CP字段 匹配每个拆分结果
            cpList = cpVal.split(';')
            for cpItem in cpList:
                match = pattern.match(cpItem)
                if match:
                    TmpArray[CP_Index] = match.group().split(' ')
                    row_index = Excel_WriteLine(out_table, row_index, TmpArray)

        else:
            row_index = Excel_WriteLine(out_table, row_index, lineArray)

    # 保存文件
    out_file.save(output_path)
    print("文件生成完成")





if __name__ == '__main__':
    Excel_Process()


