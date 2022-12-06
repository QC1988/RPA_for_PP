#!/usr/bin/env python
# -*- coding: <utf-8> -*-

import xlwings as xw
import time
import datetime as dt
from logging import getLogger, StreamHandler,FileHandler,Formatter,DEBUG,ERROR

# log関係設定
formatter = Formatter('[%(levelname)s] %(asctime)s - %(message)s (%(filename)s)')
logger = getLogger(__name__)

# handler = StreamHandler()
handler = FileHandler("log.txt")
handler.setLevel(DEBUG)
handler.setFormatter(formatter)

error_handler = FileHandler("error.txt")
error_handler.setLevel(ERROR)
error_handler.setFormatter(formatter)

logger.setLevel(DEBUG)
logger.addHandler(handler)
logger.addHandler(error_handler)

# 何行目からデータ
row_num_start = 8
# 何行目までコピーする　3000でOK
row_num_max = 3000
count = 0
# V列 j = 22 / Z列 j = 26 / AA列 j = 27 / AH列 j = 34
column_need_to_copy = (22, 26, 27, 34)
# M列　J4P　13　/　P列  仕入先　16　/　P列　16 NJRC,APD,FIT,ｻﾝｹﾝ,ﾕﾀｶ
checker = 13
selector = 16
supplier = ("NJRC","APD","FIT","ｻﾝｹﾝ","ﾕﾀｶ")
# M列　J49　Noで検証する　コピーとオリジナルの行が一致するのか、ずれるとエラーを出す
check_code = 13

#　数字をアルファベットに変換する　1-->'A' 27-->'AA'  入力値範囲1~51
def changeNumToChar(toBigChar):
    increment = 0
    res_char = ''
    increment = ord('A') - 1
    if toBigChar <= 26:
        res_char = chr(toBigChar+increment)
    elif toBigChar > 26:
        shang,yu = divmod(toBigChar, 26)
        char = chr(yu + increment)
        res_char = chr(shang + increment) + char
    return res_char


# メッセージを表示する
print("================copy ups_scheduler_copy to original==============")
print("=       Ver.1.0                                                 =")
print("=       2022/10/24                                              =")
print("=       IoTG QC                                                 =")
print("=  1.UPS日程表をコピーする                                      =")
print("=  UPS日程表 201106 のコピー.xlsm →　UPS日程表 201106 .xlsm    =")
print("=  V列、Z列、AA列、AH列　{0}行目~{1} 文字とセルの色              =".format(row_num_start,row_num_max))
print("=================================================================")
print("")

# Excel窓口を表示する
app = xw.App(visible=True)

# メッセージを表示する
print("Excel will be open.Plesase wait...")

# 現在日付時刻を取得する
localtime = time.localtime(time.time())
# 年月日時分秒の順で表示する
year_month_day_hour_min_sec = dt.datetime(localtime.tm_year, localtime.tm_mon, localtime.tm_mday,localtime.tm_hour,localtime.tm_min,localtime.tm_sec)
print(year_month_day_hour_min_sec)
logger.info(year_month_day_hour_min_sec)

# Excelファイルを開き、読み取り専用アラームにNoを選択する　コピー版
wb_copy = app.books.open(fullname='C:\\Users\\035203557\\Desktop\\kaizen_space\\RPA\\8.UPS_Update_production_schedule\\UPS日程表 201106  - コピー.xlsm',ignore_read_only_recommended=True)  # on Windows: use raw strings to escape backslashes
# 1番目のsheetを選択する
sht_copy = wb_copy.sheets[0]

# Excelファイルを開き、読み取り専用アラームにNoを選択する　オリジナル版 
wb_original = app.books.open(fullname='C:\\Users\\035203557\\Desktop\\kaizen_space\\RPA\\8.UPS_Update_production_schedule\\UPS日程表 201106 .xlsm',ignore_read_only_recommended=True)  # on Windows: use raw strings to escape backslashes
# 1番目のsheetを選択する
sht_original = wb_original.sheets[0]

# Excelにデータある行の番号
row_num_max = sht_copy.used_range.rows.count + 1
print(row_num_max)

# メッセージを表示する
print("Excel have been opened.")

# 列 j
for j in column_need_to_copy:
    #　行 i
    for i in range(row_num_start, row_num_max):
        # UPS_copyとUPS_original同じ行かをチェックする　M列だけ→（M列+N列）　J4P+分割No.　重複しているJ4Pがあるため
        # 同じ行の場合、そのままコピーする　
        # 違う場合、前100行から後ろ100行まで（M列+N列）で合ってるかどうかを確認する
        if sht_copy.range((i,checker)).value == sht_original.range((i,checker)).value:
            if sht_copy.range((i,selector)).value in supplier:
                UPS_copy_Value = sht_copy.range((i,j)).value
                UPS_copy_color = sht_copy.range((i,j)).color
                UPS_original_Value = sht_original.range((i,j)).value
                UPS_originai_color = sht_original.range((i,j)).color
                if UPS_copy_Value == UPS_original_Value:
                    pass
                    # print("セル{0}{1}変更✕".format(changeNumToChar(j),i),end=" ")
                else:
                    print("%s"%(sht_copy.range((i,selector)).value),end="   ")
                    print("{0}{1}   値変更◯".format(changeNumToChar(j),i),end="    ")
                    print("{}→{}".format(UPS_original_Value,UPS_copy_Value))
                    logger.info("%s"%(sht_copy.range((i,selector)).value))
                    logger.info("{0}{1}   値変更◯".format(changeNumToChar(j),i))
                    logger.info("{}→{}".format(UPS_original_Value,UPS_copy_Value))
                    sht_original.range((i,j)).value = UPS_copy_Value
                    count = count + 1

                if UPS_copy_color == UPS_originai_color:
                    # print("セル{0}{1}変更✕".format(changeNumToChar(j),i))
                    pass
                else:
                    print("%s"%(sht_copy.range((i,selector)).value),end="   ")
                    print("{0}{1}   色変更◯".format(changeNumToChar(j),i))
                    logger.info("%s"%(sht_copy.range((i,selector)).value))
                    logger.info("{0}{1}   色変更◯".format(changeNumToChar(j),i))
                    sht_original.range((i,j)).color = UPS_copy_color
                    count = count + 1

            else:
                # print("仕入先:%s"%(sht_copy.range((i,selector)).value),end=" ")
                # print("対象外、コピー無し")
                pass
        else:
            print("エラー！UPS日程表とコピー版の行が異なります！")
            logger.error("エラー！UPS日程表とコピー版の行が異なります！")
print("変更箇所総数：%i"%count)
logger.info("変更箇所総数：%i"%count)
logger.info("以上")
logger.info("")

# 5秒待機する
time.sleep(5)

# Excelファイル保存する
wb_original.save()

# 開いたExcelファイルを閉じる
wb_copy.close()
wb_original.close()

# Book1を閉じ、Excel appを閉じる
app.quit()

# メッセージを表示する
print("Excel has been closed.")
