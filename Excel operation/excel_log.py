#!/usr/bin/env python
# -*- coding: <utf-8> -*-

from logging import getLogger, StreamHandler,FileHandler,Formatter,DEBUG,ERROR

# メッセージを表示する
print("================write log in log.txt and error.txt===============")
print("=       Ver.1.0                                                 =")
print("=       2022/10/24                                              =")
print("=       IoTG QC                                                 =")
print("=  1.log.txtに記載する                                          =")
print("=  2.error.txtに記載する                                        =")
print("=================================================================")
print("")

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

logger.debug("これはデバッグです")
logger.info("プログラムが開始しました")
logger.warning("ファイルの容量が200GBを超えました")
logger.error("ファイルが存在していません")
logger.critical("重大問題発生しました")
