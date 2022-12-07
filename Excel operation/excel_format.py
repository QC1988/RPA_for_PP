#!/usr/bin/env python
# -*- coding: <utf-8> -*-

import xlwings as xw
import time


# メッセージを表示する
print("================open excel and close with xlwings================")
print("=       Ver.1.0                                                 =")
print("=       2022/10/21                                              =")
print("=       IoTG QC                                                 =")
print("=  1.セル背景色設定                                             =")
print("=  2.フォント設定                                               =")
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

# range範囲を選択する
rng = sht.range("B2")

# セル背景色の設定方法　255,255,000　黄
rng.color = 255,255,000
# セル背景色の設定方法　255,192,000　黄
rng.color = 255,192,000
# セル背景色の設定方法　153,255,51　緑
rng.color = 153,255,51

# セルの背景色の設定方法　
sht.range("B2").color = (153,255,51)

# セル背景色を取得する
clr_B2 = sht.range("B2").color
print(clr_B2)

# セル背景色をクリアする
sht.range("B2").color = None

# セルに数値を入力する
rng.value = "test1test2test3"

# フォントの設定方法
## 文字色　青
rng.api.Font.ColorIndex = 5
## 文字色　赤
rng.api.Font.ColorIndex = 3
## 文字色　黄
rng.api.Font.ColorIndex = 6
## フォントサイズ
rng.api.Font.Size = 30
## 太字
rng.api.Font.Bold = True
## 横　中央揃え
rng.api.HorizontalAlignment = -4131
## 縦　中央揃え
rng.api.VerticalAlignment = -4130
## セル幅を自動調整
rng.autofit()

# 5秒待機する
time.sleep(10)

# Excelファイル保存する
wb_source.save()

# 開いたExcelファイルを閉じる
wb_source.close()

# Book1を閉じ、Excel appを閉じる
app.quit()

# メッセージを表示する
print("Excel has been closed.")
