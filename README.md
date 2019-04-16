# python_xlsx2csv
エクセルファイルからcsvファイルへ、条件つきで変換する。

# 使用ライブラリ
xlrd

## 機能
- エンコード、改行コードを指定できる。
- スキップする行を指定できる。
- 文字列か数値か判断して、出力形式を分ける。
- 小数(パーセント)は桁数指定

## ロジック詳細

### import

```
import xlrd
```

### エクセル読み取り
```
xls = xlrd.open_workbook(INPUT_XLSX)
sheet = xls.sheet_by_index(0)

rows = sheet.nrows # 行数 
cols = sheet.ncols # 列数

for r in range(0, rows):
    for c in range(0, cols):
        cell = sheet.cell(r,c)
```

```
cell.value # 値を取得
cell.ctype # 値のタイプを取得
```

##  CSV出力
### エンコード、改行コード

```
file = open(OUTPUT_CSV, 'w', encoding = ENCODE, newline = NEWLINE)
```
```
ENCODE =>
'utf_8_sig' # UTF8 BOM付き
'utf_8' # UTF8
```
```
NEWLINE =>
'\r\n' # CRLF
'\n' # LF
```

### 出力用関数

```
# 文字列出力 (""で囲む)
def writeStr(value):
    return '\"' + str(value) + '\"'

# 数値出力
def writeNum(value):
    return str(value)
```

### 文字列か否か

```
# 文字列の場合
if(cell.ctype == xlrd.XL_CELL_TEXT):
    writeStr(cell.value)
```

### 数値か否か

```
if(cell.ctype == xlrd.XL_CELL_NUMBER):
    if(cell.value.is_integer()):
      writeNum(int(cell.value))
    # 少数の時 第3位まで
    else:
      writeNum(round(cell.value, ROUND))        
```
```
# 許有少数桁
ROUND = 3
```

### 整数か否か

```
value.is_integer()
```

### 実施方法

```
python3 tocsv.py
```
