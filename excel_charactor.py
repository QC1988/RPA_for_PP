#!/usr/bin/env python
# -*- coding: <utf-8> -*-

import xlwings as xw
import time


# メッセージを表示する
print("================open excel and close with xlwings================")
print("=       Ver.1.0                                                 =")
print("=       2022/10/21                                              =")
print("=       IoTG QC                                                 =")
print("=  1.値を入力する                                               =")
print("=  2.値を取得する                                               =")
print("=================================================================")
print("")

# Excel窓口を表示する
app = xw.App(visible=True)

# メッセージを表示する
print("Excel will be open.")

# Excelファイルを開き、読み取り専用アラームにNoを選択する
wb_source = app.books.open(fullname='',ignore_read_only_recommended=True)  # on Windows: use raw strings to escape backslashes

# sheetを選択する
sht = wb_source.sheets[0]


# Excel最後の列の番号を取得する
# sht.cells.last_cell.column
# Excel最後の行の番号を取得する
# sht.cells.last_cell.row


# Excelにデータある行の番号と列の番号
# nrow = sht.used_range.rows.count
# ncol = sht.used_range.columns.count

# セルに数値を入力する　A1セルに「1000」
sht.range("A1").value = 1000

# 内容をクリアし、空白のセルにする
sht.range("A1").clear_contents()

# セルに数値を入力する　A2セルに「2000」
sht.range((2,1)).value = 2000

# セルに数値を入力する　A3セルに「3000」Output：3000.0
sht.cells(3,1).value = 3000
# デフォルトでは数字のセルは float として読み込まれますが、int に変更できる　Output：3000
sht.range('A3').options(numbers=int).value

# セルに文字列を入力する　A4セルに「商品名」
sht.range("A4").value = "商品名"

# セルの内容とフォーマットをクリアし、初期状態にする
sht.range("A4").clear()

# セルに算式を入力する
sht.range("A5").formula = "=sum(A1:A4)"

# A1~A4セルの値を取得する
A1 = sht.range("A1").value
print(A1)
A2 = sht.range("A2").value
print(A2)
A3 = sht.range("A3").value
print(A3)
A4 = sht.range("A4").value
print(A4)
A5 = sht.range("A5").value
print(A5)

# 5秒待機する
time.sleep(30)

# Excelファイル保存する
# wb_source.save()

# 開いたExcelファイルを閉じる
wb_source.close()

# Book1を閉じ、Excel appを閉じる
app.quit()

# メッセージを表示する
print("Excel has been closed.")
