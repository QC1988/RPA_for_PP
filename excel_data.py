#!/usr/bin/env python
# -*- coding: <utf-8> -*-

import xlwings as xw
import time
import datetime as dt

# メッセージを表示する
print("================open excel and close with xlwings================")
print("=       Ver.1.0                                                 =")
print("=       2022/10/21                                              =")
print("=       IoTG QC                                                 =")
print("=  1.本日の日付を取得する                                        =")
print("=  2.日付を計算する                                              =")
print("=================================================================")
print("")

# Excel窓口を表示する
app = xw.App(visible=True)

# メッセージを表示する
print("Excel will be open.")

# Excelファイルを開き、読み取り専用アラームにNoを選択する
wb_copy = app.books.open(fullname='C:\\Users\\035203557\\Desktop\\kaizen_space\\RPA\\8.UPS_Update_production_schedule\\UPS日程表 201106 .xlsm',ignore_read_only_recommended=True)  # on Windows: use raw strings to escape backslashes

# 本日の年月日を取得する
localtime = time.localtime(time.time())
print(localtime)

sht = wb_copy.sheets[0]
sht.range("A1").value = dt.datetime(localtime.tm_year, localtime.tm_mon, localtime.tm_mday)
A1_value = sht.range("A1").value
print("A1値:",end="")
print(A1_value) # A1値:2022-10-25 00:00:00
print("A1タイプ:",end="")
print(type(A1_value)) # A1タイプ:<class 'datetime.datetime'>

# 日付を入力すると、Excelでは自動で日付として認識される
# また、日付を読み込むと自動でdatetime型に型変換してくれる
sht.range("A2").value = "2022/10/24"
A2_value = sht.range("A2").value
print("A2値:",end="")
print(A2_value) # A2値:2022-10-24 00:00:00
print("A2タイプ:",end="")
print(type(A2_value)) # A2タイプ<class 'datetime.datetime'>

sht.range("A3").value = "2022/10/25"
A3_value = sht.range("A3").value
print("A3値:",end="") 
print(A3_value) # A3値:2022-10-25 00:00:00
print("A3タイプ:",end="")
print(type(A3_value)) # A3タイプ<class 'datetime.datetime'>

# 日付時刻の加減算を行うためには日付と時刻の情報を持つ datetime オブジェクト (datetime.datetime) とする必要がる
differ = sht.range("A1").value - sht.range("A2").value
print("A1-A2値:",end="")
print(differ) # A1-A2値:1 day, 0:00:00
print("A1-A2タイプ:",end="")
print(type(differ)) # A1-A2タイプ:<class 'datetime.timedelta'>

# 日付の差を表示する
sht.range("A4").value = differ.days
A4_value = sht.range("A4").value
print("A4値:",end="") 
print(A4_value) # A4値:1.0
print("A4タイプ:",end="")
print(type(A4_value)) # A4タイプ:<class 'float'>

sht.range("A5").value =  differ + A3_value
plus_differ = differ + A3_value
print("A1-A2値+A3:",end="")
print(plus_differ) # A1-A2値+A3:2022-10-26 00:00:00
print("A1-A2値+A3タイプ:",end="")
print(type(plus_differ)) # A1-A2値+A3タイプ:<class 'datetime.datetime'>

W8_value = sht.range("W8").value
V8_value = sht.range("V8").value
excel_differ = W8_value - V8_value
# sht.range("A6").value = excel_differ
print("W8値:",end="")
print(W8_value) # W8値:2022-11-08 00:00:00
print("V8値:",end="")
print(V8_value) # V8値:2022-11-07 00:00:00
print("W8-V8値:",end="")
print(excel_differ) # W8-V8値:1 day, 0:00:00
print("W8-V8タイプ:",end="")
print(type(excel_differ)) # W8-V8タイプ:<class 'datetime.timedelta'>

# excel内計算式で出す2つの日付の差
excel_differ_formula =  sht.range("AN9").value
print("excel内計算式で出す2つの日付の差:",end="")
print(excel_differ_formula) # 10.0
print("excel内計算式で出す2つの日付の差のタイプ:")
print(type(excel_differ_formula)) # <class 'float'>

if sht.range("A1").value == sht.range("A3").value:
    print("A1とA3が同じタイプ。") # A1とA3が同じタイプ。
else:
    print("A1とA3が異なるタイプ。")


if A4_value == 1 :
    print("日付の差がfloat型の1.0ため、int型の1と一致しない")
else:
    print("日付の差がfloat型の1.0だが、int型の1と一致する") # 日付の差がfloat型の1.0ため、int型の1と一致しない

if A4_value == 1.0:
    print("日付の差がfloat型の1.0ため、float型の1.0と一致する") # "日付の差がfloat型の1.0ため、float型の1.0と一致する"
else:
    print("日付の差がfloat型の1.0だが、float型の1.0と一致しない")   


# 5秒待機する
time.sleep(20)

# Excelファイル保存する
wb_copy.save()

# 開いたExcelファイルを閉じる
wb_copy.close()

# Book1を閉じ、Excel appを閉じる
app.quit()

# メッセージを表示する
print("Excel has been closed.")
