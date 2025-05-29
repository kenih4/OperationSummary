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
#        input("正常終了:計画時間.xlsxを開きます。手動で修正する必要がある場合は、修正してからマクロ「cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI()」を実行しましょう。\n手動で修正する必要がある例は、FCBTがある時や、BL-studyが紛れ込んでる時などです。\nEnter押すと計画時間.xlsxを開きますが、最前面にでてこないのです。。。\nPress Enter to continue...")
        input("正常に「計画時間.xlsx」が作成されました。手動で修正する必要があります。\n手動で修正する必要がある例は、FCBTがある時や、BL-studyが紛れ込んでる時などです。\n計画時間.xlsxのチェックに進んで下さい。\nPress Enter to Exit...")
#        import win32com.client                                          #Win32comモジュールを呼び出す
#        try:
#            excelapp = win32com.client.Dispatch('Excel.Application')        #Excelアプリケーションを起動する
#            excelapp.Visible = 1                                            #Excelウインドウを表示する
#            excelapp.Workbooks.Open(schedule.計画時間ファイル,ReadOnly=False)  #           
##            excelapp.Workbooks.Open(r"C:\Users\kenichi\Dropbox\gitdir\VBA運転集計用\マクロいろいろ.xlsm",ReadOnly=False)  #  マクロcp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEIは、マクロいろいろ.xlsmnのボタンで行うようにしたのでこれは不要になった
##            excelapp.Application.Run('Module2.cp_paste_KEIKAKUZIKAN_UNTENZYOKYOSYUKEI')                       #標準モジュールModule1のマクロtest1を実行する
##            excelapp.Workbooks(1).Close(SaveChanges=False)    
##            excelapp.Application.Quit()                                     #Excelを閉じる
#        finally:
#            pass
#            excelapp.Application.Quit()
    else:
        print(f"異常：作成されたはずの計画時間.xlsxのタイムスタンプが古いです。 最終更新時刻: {datetime.datetime.fromtimestamp(os.path.getmtime(schedule.計画時間ファイル))}")

#    print("schedule.計画時間ファイル:",schedule.計画時間ファイル)
#    EXCEL = schedule.計画時間ファイル
#    subprocess.Popen(['start', EXCEL], shell=True)
else:
    print("異常終了しました")
#time.sleep(60)