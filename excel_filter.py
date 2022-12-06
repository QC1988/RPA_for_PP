#!/usr/bin/env python
# -*- coding: <utf-8> -*-

import xlwings as xw
import time


# メッセージを表示する
print("================filter usage================")
print("=       Ver.1.0                                                 =")
print("=       2022/10/21                                              =")
print("=       IoTG QC                                                 =")
print("=  1.フィルターをかける                                         =")
print("=  2.フィルターを解除する                                       =")
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


# 範囲を選択する
rng = wb_source.sheets[0].range("A1:BZ3000")

# フィルターをかける 1つの列に、1つの条件を設定
rng.api.AutoFilter(Field=2, Criteria1="小")

# フィルターをかける 1つの列に、2つの条件を設定 Operator=2は、ORを表す
rng.api.AutoFilter(Field=4, Criteria1="BA100T", Operator=2, Criteria2="BL100T")

# オートフィルターをクリア
wb_source.sheets[0].api.ShowAllData()

# オートフィルターを削除 オートフィルターの設定を削除する方法です。全ての行が表示されます。
# wb_source.sheets[0].api.AutoFilterMode = False

# フィルターをかける フィルターの条件に、複数の値を指定 Operator=7は複数
rng.api.AutoFilter(Field=4, Criteria1=["BA75T","BL75T"], Operator=7)

# 5秒待機する
time.sleep(15)

# Excelファイル保存する
# wb_source.save()


# 開いたExcelファイルを閉じる
wb_source.close()

# Book1を閉じ、Excel appを閉じる
app.quit()

# メッセージを表示する
print("Excel has been closed.")
