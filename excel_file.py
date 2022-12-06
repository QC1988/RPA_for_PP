#!/usr/bin/env python
# -*- coding: <utf-8> -*-

import xlwings as xw
import time


# メッセージを表示する
print("================open excel and close with xlwings================")
print("=       Ver.1.0                                                 =")
print("=       2022/10/21                                              =")
print("=       IoTG QC                                                 =")
print("=  1.Excelを開く                                                =")
print("=  2.Excelを閉じる                                              =")
print("=================================================================")
print("")

# Excel窓口を表示する
app = xw.App(visible=True)

# メッセージを表示する
print("Excel will be open.")

# this will create a new workbook
# wb = xw.Book()  
# connect to a file that is open or in the current working directory
# wb = xw.Book('FileName.xlsx') 

# Excelファイルを開き、読み取り専用アラームにNoを選択する
wb_source = app.books.open(fullname='C:\\Users\\035203557\\Desktop\\納期調整\\テスト用.xlsm',ignore_read_only_recommended=True)  # on Windows: use raw strings to escape backslashes

# 5秒待機する
time.sleep(5)

# Excelファイル保存する
wb_source.save()

# 開いたExcelファイルを閉じる
wb_source.close()

# Book1を閉じ、Excel appを閉じる
app.quit()

# メッセージを表示する
print("Excel has been closed.")
