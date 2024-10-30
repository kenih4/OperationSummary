#!/usr/bin/python
# -*- coding: utf-8 -*-

from mdaq import pymdaq_web
import datetime
import time
import pandas as pd
import sys
import codecs
import os
import sys
import binascii
import requests
from requests.exceptions import Timeout
import pandas as pd
import openpyxl
import subprocess
import copy
import libCom
import schedule
import GunHvOff






ret = GunHvOff.output_excel_gun_hvoff_time()
if ret == 0:
#    print("正常終了しました")
    
    if abs(time.time() - os.path.getmtime(schedule.計画時間ファイル))<10:
        print("正常終了:マクロいろいろ.xlsmが立ち上がるので、マクロ「cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI()」が実行されます")
        #subprocess.Popen(['start', r"C:\Users\kenichi\Documents\OperationSummary\マクロいろいろ.xlsm"], shell=True) #マクロが入ってるエクセルファイルを開く
        import win32com.client                                          #Win32comモジュールを呼び出す
        try:
            excelapp = win32com.client.Dispatch('Excel.Application')        #Excelアプリケーションを起動する
            excelapp.Visible = 1                                            #Excelウインドウを表示する
            excelapp.Workbooks.Open(r"C:\Users\kenichi\Documents\OperationSummary\マクロいろいろ.xlsm",ReadOnly=True)  #rを追加してパス名をrawデータとして読み込みマクロ有効ブックを開く
            excelapp.Application.Run('Module2.cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI')                       #標準モジュールModule1のマクロtest1を実行する
#            excelapp.Workbooks(1).Close(SaveChanges=False)    
#            excelapp.Application.Quit()                                     #Excelを閉じる
        finally:
            pass
#            excelapp.Application.Quit()  

    else:
        print(f"異常：作成されたはずの計画時間.xlsxのタイムスタンプが古いです。 最終更新時刻: {datetime.datetime.fromtimestamp(os.path.getmtime(schedule.計画時間ファイル))}")

    #print("schedule.計画時間ファイル:",schedule.計画時間ファイル)
    #EXCEL = schedule.計画時間ファイル
    #subprocess.Popen(['start', EXCEL], shell=True)
else:
    print("異常終了しました")
time.sleep(60)