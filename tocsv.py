# -*- coding: utf-8 -*-
import xlrd

# 入力エクセル
INPUT_XLSX = 'input.xlsx'
# 出力CSV
OUTPUT_CSV = 'output.csv'

# encode
# ENCODE = 'utf_8_sig' # UTF8 BOM付き
ENCODE = 'utf_8' # UTF8

# 改行コード
NEWLINE = '\r\n' # CR+LF
# NEWLINE = '\n' # LF    

# スキップ行(例)
SKIP1 = "タイトル"
SKIP2 = "合計"

# 許有少数桁
ROUND = 5

# 文字列出力 (""で囲む)
def writeStr(value):
    return '\"' + str(value) + '\"'

# 数値出力
def writeNum(value):
    return str(value)


xls = xlrd.open_workbook(INPUT_XLSX)
sheet = xls.sheet_by_index(0)

# UTF8 BOM付き 改行コードCRLF
file = open(OUTPUT_CSV, 'w', encoding = ENCODE, newline = NEWLINE) 

rows = sheet.nrows # 行数 
cols = sheet.ncols # 列数

for r in range(0, rows):
    sb = ''
    for c in range(0, cols):
        cell = sheet.cell(r,c)

        # スキップ行
        if(cell.value == SKIP1 or cell.value == SKIP2):
            break

        # 文字列の場合
        if(cell.ctype == xlrd.XL_CELL_TEXT):
            sb += writeStr(cell.value)

        # 数値の場合
        elif(cell.ctype == xlrd.XL_CELL_NUMBER):
            if(cell.value.is_integer()):
                sb += writeNum(int(cell.value))
            # 少数の時第3位まで
            else:
                sb += writeNum(round(cell.value, ROUND))               
        if(c != cols-1):
            sb += ','
        # 最終列は改行、書き込み
        else:
            sb += '\n'
            file.write(sb)

file.close()