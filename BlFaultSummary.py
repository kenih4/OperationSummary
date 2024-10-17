#!/usr/bin/env python
# coding: utf-8

# In[1]:


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
from operator import itemgetter
import re


# In[ ]:





# In[2]:


#OFF/ONリストの合計時間を表示する
def get_total_time(list):
    total_sec = 0.0
    for i in range(len(list)):
        dSec = (list[i]["end"] - list[i]["start"]).total_seconds()
        list[i]["total"] = disp_time(dSec)
        total_sec = total_sec + dSec
    return list,str(len(list)), disp_time(total_sec)   
    
#"hh:mm:ss"形式で表示する     
def disp_time(total_sec):
    td = datetime.timedelta(seconds=total_sec)
    m, s = divmod(td.seconds, 60)
    h, m = divmod(m, 60)
    timeStr = str(h).zfill(2) +":" + str(m).zfill(2) +":" + str(s).zfill(2)
    return timeStr

#合計時間を分単位で四捨五入する
def disp_time_floor_minute(total_sec):
    offset = 0
    td = datetime.timedelta(seconds=total_sec)
    m, s = divmod(td.seconds, 60)
    h, m = divmod(m, 60)
    
    if s >30:
        offset = 1
    timeStr = str(h).zfill(2) +":" + str(m+offset).zfill(2)  +":" + str(0).zfill(2)
    return timeStr


# In[3]:


#リストから他のBLのFaultを消す
def get_list_bl_select(list, bl_name):
    list_tmp = []
    BL = bl_name.upper()
        
    for i in range(len(list)):
        if list[i]["BL"]  == "ALL" or list[i]["BL"]  == BL:
            list_tmp.append(list[i]) 
    return list_tmp


# In[4]:


#ユニットのリストをBL集計用の形式でtxtに出力する    
def output_log_txt(list):
    txt = 'fault.txt'
    with open(txt,'w') as f:
        for i in range(len(list)):
            f.write( "\t".join(list[i]) + "\n")
    subprocess.Popen(['start', txt], shell=True)

#ユニットのリストをBL集計用の形式でtxtに出力する  「時間指定する」
def output_log_txt_Time_Specification(list,dt_beg,dt_end):
    import pyperclip
    txt = 'fault.txt'
    txt_clibbord = ''
    with open(txt,'w', encoding='utf-8') as f:  #なぜか、if dt > dt_beg　の条件をいれると文字化けするので encoding='utf-8'と指定した
        for i in range(len(list)):
            dt = datetime.datetime.strptime(list[i][2], '%Y/%m/%d %H:%M:%S')
            if dt > dt_beg and dt < dt_end:
                f.write( "\t".join(list[i]) + "\n")
                txt_clibbord += "\t".join(list[i]) + "\n"
    if txt_clibbord == '':
        print("txt_clibbordが空っぽです。前のステップで、SACLA運転状況集計BL*.xlsmを保存していない可能性があります。次ぎ進んでもしょうがないのでここで終了します。")
        sys.exit()
    pyperclip.copy(txt_clibbord)
#    subprocess.Popen(['start', txt], shell=True)

# In[5]:


#ユニットのリストをBL集計用の形式でListに出力する   
def output_bl_fault_List(list, start_time, end_time):
    total_sec = 0.0
    tuning_sec = 0.0
    trip_count = 0
    output_line_list =[]
    reason_list = []
    fault_List = []
    shift_fault_list = []
        
    for i in range(len(list)):
        dSec = (list[i]["end"] - list[i]["start"]).total_seconds()
        if list[i]["運転種別"] == "調整時間":
            fault = ""
            tuning_sec += dSec
        else :
            fault = "RF"
            trip_count += 1
            
        reason_list.extend(list[i]["調整理由"]) 
        reason =  " , ".join(filter(None, set(list[i]["調整理由"])))

        
            
        output_line_list = [list[i]["BL"], list[i]["運転種別"], list[i]["start"].strftime("%Y/%m/%d %H:%M:%S"), list[i]["end"].strftime("%Y/%m/%d %H:%M:%S"), fault ,  reason]
        fault_List.append(output_line_list)
        total_sec += dSec       
    
    if len(fault_List) == 0:
        fault_List.append(["","","","","",""])   
    reason_all =  " , ".join(filter(None, set(reason_list)))
    output_line_list = ["", "ユーザー運転", start_time.strftime("%Y/%m/%d %H:%M:%S"), end_time.strftime("%Y/%m/%d %H:%M:%S"), str(trip_count) , reason_all]
    fault_List.append(output_line_list)
    
    shift_fault_list = get_unit_fault_List(start_time, end_time, tuning_sec, total_sec, trip_count)
    

    return fault_List, shift_fault_list


#ユニット集計用のタブ形式でListに出力する  
def get_unit_fault_List(start_time, end_time, tuning_sec, total_sec, trip_count):
    unit_fault_list = []
    
    shift_sec = (end_time-start_time).total_seconds()
    shift_time = disp_time_floor_minute(shift_sec)
    utility_time = disp_time_floor_minute(shift_sec-total_sec)
    rate = (shift_sec-total_sec)/shift_sec
    fault_time = disp_time_floor_minute(total_sec - tuning_sec)
    fault_interval = disp_time_floor_minute(shift_sec/(trip_count+1))
    
    unit_fault_list = [start_time.strftime("%Y/%m/%d %H:%M:%S"), end_time.strftime("%Y/%m/%d %H:%M:%S"), shift_time, utility_time, rate, 
                      disp_time_floor_minute(tuning_sec), fault_time, disp_time_floor_minute(total_sec), trip_count, fault_interval]
    
    return unit_fault_list
   


# In[6]:


#期間をシフト単位に区切ったList取得する。
def get_user_period_time_list(start_time, end_time):
    time_list = []
    period_list = []

    
    if (end_time-start_time).days < 0:
        ret = -1
        print("Time ERROR")
    else:
        time_bk = start_time
        while (time_bk - end_time).days < 0:
            time_list.append(time_bk)
            time_bk = time_bk + datetime.timedelta(hours = 12)

        time_list.append(end_time)
        
        for i in range(len(time_list)):
            if i < (len(time_list)-1):
                 period_list.append([time_list[i], time_list[i+1]])
    
    return period_list


# In[7]:


#Fault記録用のExcelからTrip時間と調整時間を取得したlistを返す /重複削除の処理はしていない/
def get_fault_list(acc):    
    fault_list = schedule.read_xcel_fault_time(acc, "停止時間")
    tuning_list = schedule.read_xcel_fault_time(acc, "調整時間")
    list = fault_list + tuning_list
    sorted_list = sorted(list, key=itemgetter("start", "end"))
    return sorted_list 


# In[8]:


#リストの重複する時間をまとめる
def merge_periods_list(list):
    list_bk =list.copy()
    list_tmp = []
    dict = {}
    
    for i in range(len(list_bk)):
        dict = list_bk[i].copy() #start time
        dict.update({"flg":1, "time":list_bk[i]["start"]})
        list_tmp.append(dict)
        dict = {}
        dict = list_bk[i].copy() #start time
        dict.update({"flg":-1, "time":list_bk[i]["end"]})
        list_tmp.append(dict)
        
    sorted_list = sorted(list_tmp, key=itemgetter('time', 'flg'))

    total = 0
    for i in range(len(sorted_list)):
        total += sorted_list[i]["flg"]
        sorted_list[i]["Sum"] = total 

    Sum_bk = 0
    offOnList = []
    reason_tmp = []
    opr_kind_tmp =  ""
    BL = ""
    
    
    for i in range(len(sorted_list)):
        if(sorted_list[i]["調整理由"] != ""):
            reason_tmp.append(sorted_list[i]["調整理由"])
        if(sorted_list[i]["運転種別"] == "調整時間"):
            opr_kind_tmp = "調整時間"
        if(sorted_list[i]["BL"] != "ALL"):
            BL = sorted_list[i]["BL"]

            
        if(sorted_list[i]["Sum"] == 1 and Sum_bk == 0):
            off_time = sorted_list[i]["time"]
        if(sorted_list[i]["Sum"] == 0 and Sum_bk == 1):
            on_time = sorted_list[i]["time"]
            offOnList.append({'運転種別': opr_kind_tmp, 'BL': BL, 'start': off_time, 'end': on_time, '調整理由': reason_tmp})
            reason_tmp = []
            opr_kind_tmp = ""
            BL = ""
           
        Sum_bk = sorted_list[i]["Sum"]

    return  offOnList


# In[9]:


def get_shift_list(fault_list, start_time, end_time, bl_name):
    period_list = schedule.get_list_period_time(fault_list,  start_time, end_time) #期間を抽出したリストを返す
    bl_list = get_list_bl_select(period_list, bl_name) #対象のBLの時間を抽出したリストを返す        
    merge_list = merge_periods_list(bl_list) #重複時間を削除したリストを返す。
    edit_fault_list = edit_fault_list_time(merge_list, start_time, end_time)
    output_list,shift_list = output_bl_fault_List(edit_fault_list, start_time, end_time) #リストを出力用の形式に変換   
    return output_list,shift_list


# In[10]:


def get_user_shift_time_list(bl_num):
    user_shift_time_list = []
    
    
    #計画時間のユーザー時間のみのリストを取得
    user_list = schedule.read_xcel_bl_operation_time_2(bl_num)
    user_list = schedule.extract_list_specified_key(user_list, '運転種別', 'ユーザー', 0)

    for i in range(len(user_list)):        
        user_shift_time_list.extend(get_user_period_time_list(user_list[i]["start"], user_list[i]["end"]))
    
    return user_shift_time_list
    


# In[11]:


def get_unit_list(bl_name, fault_list):
    unit_fault_list = []
    
    user_shift_time_list = get_user_shift_time_list(int(re.sub(r"\D", "", bl_name)))
    
    for i in range(len(user_shift_time_list)):
        tmp, tmp1 = get_shift_list(fault_list, user_shift_time_list[i][0], user_shift_time_list[i][1], bl_name)
        unit_fault_list.extend(tmp)
    return unit_fault_list

    


# In[12]:


#Fault listのトリップ時間を開始時間と終了時間で区切る。(期間内抽出、重複処理後に使用する)
def edit_fault_list_time(fault_list, start_time, end_time):
    fault_list_bk =  copy.copy(fault_list)  
      
    for i in range(len(fault_list_bk)):
        if (start_time <= fault_list_bk[i]["end"]) and (end_time >= fault_list_bk[i]["start"]):#期間内判定
            if start_time > fault_list_bk[i]["start"]:
                fault_list_bk[i]["start"] = start_time
            if end_time < fault_list_bk[i]["end"]:
                fault_list_bk[i]["end"] = end_time
    return fault_list_bk


# In[13]:


# bl = input("BLを選択してください。(bl1,bl2,bl3)>>>")
# if bl == "bl1":
#     acc = 1
# else:
#     acc = 0

# fault_list = get_fault_list(acc)
# list = get_unit_list(bl, fault_list)
# libCom.print_list(list)
# output_log_txt(list)






# In[ ]:





# In[ ]:




