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
from icalendar import Calendar, Event
import icalendar
import binascii
import requests
from requests.exceptions import Timeout
import pandas as pd
import openpyxl
import subprocess
from operator import itemgetter
import copy
import schedule


# In[2]:


args = sys.argv

begindate = args[1]
enddate = args[2]
acc = int(args[3]) #1:SCSS, other:SACLA
begindate = begindate.replace('+', ' ') + ":00"
enddate = enddate.replace('+', ' ') + ":00"

#test
# begindate = "2022/7/22 17:00:00"
# enddate = "2022/7/23 1:00:00"
# acc = 0
#test_end


# In[3]:


if acc == 1:#SCSS
    server = "xfweb-dmz-03" 
    sig_list = ['scss_tmg_gun_tdu/status','']
    txt_name = "log_scss.txt"
    txt_tab ="[SCSS+ BL1]"
    EXCEL = schedule.SCSS集計ファイル
else:#SACLA
    server = "srweb-dmz-03" 
    sig_list =['xfel_tmg_gun_tdu/status', 'xfel_safety_oper_intlk_info2/status']
    txt_name = "log_sacla.txt"
    txt_tab ="[SACLA BL2, BL3]"
    EXCEL = schedule.SACLA集計ファイル
flgINPUT = 1
    

db = pymdaq_web.db(server, port=8888)


# In[4]:


#リストを1行ずつprintする
def print_list(list):
    for i in range(len(list)):
        print(list[i])


# In[5]:



#範囲時間内のchopperのステータスを取得
def get_chopper_status_db(acc, dt_beg, dt_end):
    chopper_list = []   
    ret = check_db_get_time(dt_beg, dt_end)
    
    if ret == 0:   
        res = db.get_data_multi(sig_list,dt_beg.strftime("%Y/%m/%d %H:%M:%S"),dt_end.strftime("%Y/%m/%d %H:%M:%S"))
        if db.status()==pymdaq_web.DB_OK:
            if acc == 1:
                chopper_list = get_chopper_off_time_soft(res['scss_tmg_gun_tdu/status']['res'], 12) #scss chopper soft                               
            else:
                chopper_list = make_sacla_chopper_list(res, flgINPUT)
            return chopper_list
        else:
            print("ERR: %d  Msg: %s"%(db.status(),db.err_msg()))

            
#時間が3日未満か確認する
def check_db_get_time(dt_beg, dt_end):

    limit_day =  3
    
    if (dt_end-dt_beg).days >= limit_day or (dt_end-dt_beg).days < 0:
        ret = -1
        print("DB GET Time ERROR")
    else:
        ret = 0    
    return ret

#ステータスリストからoff/on時間のリストを取得する
def get_chopper_off_time_soft(sig_res, bit1):
    sig_res.sort()
    list = []
    status = "ON"

    
    for i in range(len(sig_res)):
        if (sig_res[i][1] & (1 << bit1) != 0) :
            if status == "ON":
                off_time = datetime.datetime.strptime(sig_res[i][0], "%Y/%m/%d %H:%M:%S.%f").replace(microsecond = 0)
                status = "OFF"
        else:
            if status == "OFF":
                on_time = datetime.datetime.strptime(sig_res[i][0], "%Y/%m/%d %H:%M:%S.%f").replace(microsecond = 0)
                status = "ON"
                
                td = on_time - off_time
                if td.total_seconds() >2:#2秒以下を排除
                    list.append({"start":off_time, "end":on_time})
        if (status == "OFF") & (i == (len(sig_res)-1)):
            on_time = datetime.datetime.strptime(sig_res[i][0], "%Y/%m/%d %H:%M:%S.%f").replace(microsecond = 0)
            list.append({"start":off_time, "end":on_time})
    return list

#ステータスリストからBL2/BL3のoff/on時間のリストを取得する
def get_chopper_off_time_permission(sig_res, bit1, bit2):
    sig_res.sort()
    list = []
    status = "ON"

    
    for i in range(len(sig_res)):
        if (sig_res[i][1] & (1 << bit1) == 0) & (sig_res[i][1] & (1 << bit2) == 0) :
            if status == "ON":
                off_time = datetime.datetime.strptime(sig_res[i][0], "%Y/%m/%d %H:%M:%S.%f").replace(microsecond = 0)
                status = "OFF"
        else:
            if status == "OFF":
                on_time = datetime.datetime.strptime(sig_res[i][0], "%Y/%m/%d %H:%M:%S.%f").replace(microsecond = 0)
                status = "ON"
                
                td = on_time - off_time
                if td.total_seconds() >2:#2秒以下を排除
                    list.append({"start":off_time, "end":on_time})
        if (status == "OFF") & (i == (len(sig_res)-1)):
            on_time = datetime.datetime.strptime(sig_res[i][0], "%Y/%m/%d %H:%M:%S.%f").replace(microsecond = 0)
            list.append({"start":off_time, "end":on_time})
    return list



#指定した条件のChopper OFF時間を取得する
def make_sacla_chopper_list(res, flg): # 1:soft+ bl2 & bl3, 2:soft+bl2, 3:soft+bl3, 5:soft+bl2+bl3, other:soft, 
    
    list = [] 
    soft_list = get_chopper_off_time_soft(res['xfel_tmg_gun_tdu/status']['res'], 10) #chopper soft
    bl2_list = get_chopper_off_time_permission(res['xfel_safety_oper_intlk_info2/status']['res'], 0, 0) #chopper bl2
    bl3_list = get_chopper_off_time_permission(res['xfel_safety_oper_intlk_info2/status']['res'], 1, 1) #chopper bl3
    bl2_3_list = get_chopper_off_time_permission(res['xfel_safety_oper_intlk_info2/status']['res'], 0, 1) #chopper bl2 OFF & bl3 OFF 
    
    if flg == 1:#chopper OFF :bl2 & bl3
        soft_list.extend(bl2_3_list)
        list = merge_periods_list(soft_list)
    elif flg == 2:#chopper OFF :bl2
        soft_list.extend(bl2_list)
        list = merge_periods_list(soft_list)     
    elif flg == 3:#chopper OFF :bl3
        soft_list.extend(bl3_list)
        list = merge_periods_list(soft_list)
    elif flg == 5:#chopper OFF :bl2 or bl3
        soft_list.extend(bl2_list)
        soft_list.extend(bl3_list)
        list = merge_periods_list(soft_list)
    else: #chopper OFF :soft
        list = soft_list

    
    return list


#リストの重複する時間をまとめる
def merge_periods_list(list):
    list_bk =copy.copy(list)
    list_tmp = []
    dict = {}
    
    for i in range(len(list_bk)):
        dict = {"flg":1,"time":list_bk[i]["start"]} #start time
        list_tmp.append(dict)
        dict = {"flg":-1,"time":list_bk[i]["end"]} #end time
        list_tmp.append(dict)
        
    sorted_list = sorted(list_tmp, key=itemgetter('time', 'flg'))
    total = 0
    for i in range(len(sorted_list)):
        total += sorted_list[i]["flg"]
        sorted_list[i]["Sum"] = total 

    Sum_bk = 0
    offOnList = []
    
    
    for i in range(len(sorted_list)):
        if(sorted_list[i]["Sum"] == 1 and Sum_bk == 0):
            off_time = sorted_list[i]["time"]
        if(sorted_list[i]["Sum"] == 0 and Sum_bk == 1):
            on_time = sorted_list[i]["time"]
            offOnList.append({"start":off_time, "end":on_time})
        Sum_bk = sorted_list[i]["Sum"]

    return  offOnList

#OFF/ONリストのdatetimeをtext"%Y/%m/%d %H:%M:%S"へ変換する
def datetime_list_format(list):
    list1 = []
    for i in range(len(list)):
        list1.append([list[i][0].strftime("%Y/%m/%d %H:%M:%S"), list[i][1].strftime("%Y/%m/%d %H:%M:%S")])
    return list1

#OFF/ONリストの合計時間を表示する
def get_total_time(list):
    total_sec = 0.0
    
    if len(list) == 0:
        return 0
    
    print(txt_tab)
    for i in range(len(list)):
        dSec = (list[i]["end"] - list[i]["start"]).total_seconds()
        reason = "RF Trip"
        
        if "調整理由" in list[i]:
            reason = list[i]["調整理由"]
            
        print(list[i]["start"].strftime("%Y/%m/%d %H:%M:%S") + "\t" +  list[i]["end"].strftime("%Y/%m/%d %H:%M:%S")+ "\t" + disp_time(dSec) + "\t" + reason)
        total_sec = total_sec + dSec       
    print("合計" + str(len(list)) +"回 " + disp_time(total_sec))
    

#リストをtxtに出力する    
def output_log_txt(list, txt_name):
    total_sec = 0.0
    with open(txt_name,'w') as f:
        f.write(txt_tab + " 集計記録貼り付け用\n")
        for i in range(len(list)):
            dSec = (list[i]["end"] - list[i]["start"]).total_seconds()
            f.write(list[i]["start"].strftime("%Y/%m/%d %H:%M:%S") + "\t" +  list[i]["end"].strftime("%Y/%m/%d %H:%M:%S")+ "\n")
            total_sec = total_sec + dSec       
        f.write("合計" + str(len(list)) +"回 " + disp_time(total_sec))
    
    subprocess.Popen(['start', txt_name], shell=True)

#"hh:mm:ss"形式で表示する     
def disp_time(total_sec):
    td = datetime.timedelta(seconds=total_sec)
    m, s = divmod(td.seconds, 60)
    h, m = divmod(m, 60)
    timeStr = str(h).zfill(2) +":" + str(m).zfill(2) +":" + str(s).zfill(2)
    return timeStr

   

    


# In[6]:


#期間内に含まれるのユーザー運転のList取得する。
def get_peroid_user_list(acc, start_time, end_time):

    if acc == 1:
        user_list = schedule.get_user_list(1)
    else:
        bl2_list = schedule.get_user_list(2)
        bl3_list = schedule.get_user_list(3)
        bl2_list.extend(bl3_list)
        user_list = merge_periods_list(bl2_list)
    peroid_user_list = schedule.get_list_period_time(user_list, start_time, end_time)
    
    return peroid_user_list


# In[7]:


#ユーザー運転中のChopper off timeを抽出する
def get_offtime_in_user(time_list, offonlist):
    list_tmp = []
    list = []
    
    for i in range(len(time_list)):        
        list_tmp = schedule.get_list_period_time(offonlist, time_list[i]["start"], time_list[i]["end"])
        list_tmp = edit_fault_list_time(list_tmp, time_list[i]["start"], time_list[i]["end"])        
        list.extend(list_tmp)
      
    return list



#開始時間と終了時間で区切る
def edit_fault_list_time(fault_list, start_time, end_time):
    fault_list_bk =  copy.copy(fault_list)  
      
    for i in range(len(fault_list_bk)):
        if (start_time <= fault_list_bk[i]["end"]) and (end_time >= fault_list_bk[i]["start"]):#期間内判定
            if start_time > fault_list_bk[i]["start"]:
                fault_list_bk[i]["start"] = start_time
            if end_time < fault_list_bk[i]["end"]:
                fault_list_bk[i]["end"] = end_time
    return fault_list_bk


# In[8]:


#period_list期間内のfaultに指定したkeyを追加する。
def add_tuning_reason(period_list, fault_list, add_key):
    fault_list_bk =  copy.copy(fault_list)        
    for i in range(len(period_list)):
        for j in range(len(fault_list_bk)):
            if (period_list[i]["start"] <= fault_list_bk[j]["end"]) and (period_list[i]["end"] >= fault_list_bk[j]["start"]):#期間内判定
                fault_list_bk[j][add_key] =period_list[i][add_key]
    return fault_list_bk


# In[9]:


def get_user_chopper_off_time(acc, chopper_list, start_time, end_time):#chopper OFFリストに調整時間の追加とユーザー時間の抽出を行う

    #調整時間の取得
    tuning_list = schedule.read_xcel_fault_time(acc,"調整時間")
    tuning_list = schedule.get_list_period_time(tuning_list, start_time, end_time)
    
    #調整時間とChopper OFF時間の合成
    tmp_list = tuning_list.copy()
    tmp_list.extend(chopper_list)
    tmp_list = merge_periods_list(tmp_list)
    tmp_list = add_tuning_reason(tuning_list, tmp_list, "調整理由")
    tmp_list = edit_fault_list_time(tmp_list, start_time, end_time)

    #スケジュールからユーザー運転時間の取得
    peroid_user_list = get_peroid_user_list(acc, start_time, end_time)
    #ユーザー運転中の停止時間の抽出(メール用)
    user_chopper_list = get_offtime_in_user(peroid_user_list, tmp_list)
    
    return user_chopper_list

def print_user_chopper_off_time(acc, start_time, end_time):#ユーザーchopper OFF時間をprintする
    
    chopper_list = get_chopper_status_db(acc, start_time, end_time)
    user_chopper_list = get_user_chopper_off_time(acc, chopper_list, start_time, end_time)
    
    #メール用をprint
    get_total_time(user_chopper_list)
    #集計記録貼り付け用のTXT出力
    output_log_txt(chopper_list, txt_name)
    subprocess.Popen(['start', EXCEL], shell=True)


# In[10]:


start_time = datetime.datetime.strptime(begindate, "%Y/%m/%d %H:%M:%S")
end_time = datetime.datetime.strptime(enddate, "%Y/%m/%d %H:%M:%S")

print_user_chopper_off_time(acc, start_time, end_time)









# In[ ]:




