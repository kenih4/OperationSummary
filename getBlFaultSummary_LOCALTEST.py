#!/usr/bin/python
# -*- coding: utf-8 -*-

import libCom
import schedule
import BlFaultSummary
import datetime
import sys
import os




bl = input(">BLを選択してください。(bl1,bl2,bl3) デフォルト bl3 >>> ")
if bl=="":
    bl="bl3"

if bl == "bl1":
    acc = 1
else:
    acc = 0

fault_list = BlFaultSummary.get_fault_list(acc) #   SACLA運転集計記録.xlsmからシートの集計記録を読み込む
list = BlFaultSummary.get_unit_list(bl,fault_list)
#libCom.print_list(list)
#BlFaultSummary.output_log_txt(list)    #MOTO 　時間指定なし


#以下のソースを追加 もともとあるBlFaultSummary.output_log_txtを時間指定できるようにした
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
with open("dt_beg.txt", mode='r', encoding="UTF-8") as f:
    buff_dt_beg = f.read()
print ("「fault.txt」の中身を以下の指定された範囲で切り抜いてクリップボードへコピーします。")
val = input("開始日時を入力してください。(例)2021/11/1 10:00  デフォルトは「" + str(buff_dt_beg) + "」    >>>")
if not val:
    dt_beg = datetime.datetime.strptime(buff_dt_beg, "%Y/%m/%d %H:%M")
else:
    try:
        dt_beg = datetime.datetime.strptime(val, "%Y/%m/%d %H:%M")
        with open("dt_beg.txt","w") as o:
            o.write(val)
    except ValueError:
        print ("エラー：日時のフォーマットが正しくありません。")
        sys.exit()

dt_end = dt_beg +  datetime.timedelta(days=14)
val = input("終了日時を入力してください。(例)2021/11/15 10:00   デフォルトは2週間後「" + str(dt_end) + "」です。    >>>")
if val:
    try:
        dt_end = datetime.datetime.strptime(val, "%Y/%m/%d %H:%M")
    except ValueError:
        print ("エラー：日時のフォーマットが正しくありません。")
        sys.exit()
#~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BlFaultSummary.output_log_txt_Time_Specification(list,dt_beg,dt_end)

input("正常終了:マクロいろいろ.xlsmが立ち上がるので、マクロ「cp_paste_faulttxt_UNTENZYOKYOSYUKEI()」が実行されます。 Press Enter to continue...")
import win32com.client                                          #Win32comモジュールを呼び出す
try:
    print ("bl = ",bl.replace("bl", ""))
    excelapp = win32com.client.Dispatch('Excel.Application')        #Excelアプリケーションを起動する
    excelapp.Visible = 1                                            #Excelウインドウを表示する
    excelapp.Workbooks.Open(r"C:\Users\kenichi\Documents\OperationSummary\マクロいろいろ.xlsm",ReadOnly=True)  #rを追加してパス名をrawデータとして読み込みマクロ有効ブックを開く
    excelapp.Application.Run('Module3.cp_paste_faulttxt_UNTENZYOKYOSYUKEI',bl.replace("bl", ""))                       #標準モジュールModule1のマクロtest1を実行する
#    excelapp.Workbooks(1).Close(SaveChanges=False)                  
#    excelapp.Application.Quit()                                     #Excelを閉じる
finally:
    pass
#    excelapp.Application.Quit()  









#import subprocess
#subprocess.Popen(['start', r"C:\Users\kenichi\Documents\OperationSummary\マクロいろいろ.xlsm"], shell=True) #マクロが入ってるエクセルファイルを開く

# いろいろマクロ.xlsmのマクロから開くので以下はコメントアウトした
#import schedule
#if bl == "bl1":
#    EXCEL = schedule.BL1集計ファイル
#elif bl == "bl2":
#    EXCEL = schedule.BL2集計ファイル
#elif bl == "bl3":
#    EXCEL = schedule.BL3集計ファイル
#else:
#    print("ERR!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
#subprocess.Popen(['start', EXCEL], shell=True)