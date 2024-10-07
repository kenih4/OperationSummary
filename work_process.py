#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import datetime
import numpy as np
import openpyxl as xl
from openpyxl.styles.borders import Border, Side

fname = "C:/Users/hasegawa-t/Documents/シフト/作業工程.xlsx" 
#"C:/Users/y-taj/Documents/シフト/作業工程.xlsx"

def transformation_schedule(fname = "./作業工程.xlsx",sheet_num = 0):
    dt_now = datetime.datetime.now()
    
    input_file = pd.ExcelFile(fname)
    sheet_names = input_file.sheet_names
    
    df = pd.read_excel(fname,sheet_name=sheet_names[sheet_num])
    series_month = df.iloc[0].dropna()

    delcol1 = df.columns.get_loc(series_month.index[0])
    delcol2 = df.columns.get_loc(series_month[series_month==dt_now.month].index[0])

    df = df.drop(df.columns[delcol1:delcol2], axis=1)

    series_days = df.iloc[1].dropna()

    delcol3 = df.columns.get_loc(series_days[series_days==dt_now.day].index[0])

    df.iat[0,delcol3]=dt_now.month

    df = df.drop(df.columns[delcol1:delcol3],axis=1)

    df = df.drop(df.columns[df.columns.get_loc(series_days[series_days==dt_now.day].index[0])+7*3:],axis=1)

    df = df.drop(["Unnamed: 0","Unnamed: 3","Unnamed: 4","Unnamed: 5"],axis=1).drop([2,3])

    df = df[df.drop(["Unnamed: 1","Unnamed: 2"],axis=1).isnull().all(axis=1)!=True]

    df=df[:-1]

    for i in range(2,len(df.index)):
        list = df.iloc[i].to_list()
        for j,l in enumerate(list):
            if(j>1):
                if (j-2)%3==0 :
                    if type(list[j+1]) is str:
                        list[j] = list[j]+","+list[j+1]
                    if type(list[j+2]) is str:
                        list[j] = list[j]+","+list[j+2]
                else:
                    list[j] = np.nan
        df.iloc[i]=list

    df = df.dropna(axis=1,how="all")
    list = []
    for i,value in enumerate(df.loc[1]):
        if type(value) is str:
            list.append(df.loc[1][i])
        else:
            list.append(str(int(df.loc[1][i])))
    df.iloc[1]=list
    
    df = df.rename(columns=df.iloc[1])
    df.drop([0,1],inplace=True)

    for j in range(len(df[df["作業番号"].isnull()].index)):
        list_a = df.iloc[df.index.get_loc(df[df["作業番号"].isnull()].index[j])].to_list()
        list_b = df.iloc[df.index.get_loc(df[df["作業番号"].isnull()].index[j])-1].to_list()
        for i,l in enumerate(list_a):
            if type(l) is str:
                list_b[i] += "," + l
                list_a[i]=np.nan
        df.iloc[df.index.get_loc(df[df["作業番号"].isnull()].index[j])-1]=list_b
    
    df = df[df["作業番号"].isnull()!=True]
    
    return df




def main():
    df_sr = transformation_schedule(fname,sheet_num = 0)
    df_sc = transformation_schedule(fname,sheet_num = 1)
    df = pd.merge(df_sr,df_sc,how='outer')
   # df.to_excel("C:/Users/y-taj/Documents/シフト/作業工程_1W.xlsx",index=False,header=True)
    df.to_excel("C:/Users/hasegawa-t/Documents/シフト/作業工程_1W.xlsx",index=False,header=True)
   # wb = xl.load_workbook('C:/Users/y-taj/Documents/シフト/作業工程_1W.xlsx')
    wb = xl.load_workbook('C:/Users/hasegawa-t/Documents/シフト/作業工程_1W.xlsx')
    ws = wb.active
    side1 = Side(style='thin', color='000000')
    border_aro = Border(top=side1, bottom=side1, left=side1, right=side1)
    #列幅の設定、罫線の設定
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            cell.border = border_aro
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))

        adjusted_width = (max_length + 2) * 1.8
        ws.column_dimensions[column].width = adjusted_width

    #ウインドウ枠の固定
    ws.freeze_panes = 'B1'
   # wb.save("C:/Users/y-taj/Documents/シフト/作業工程_1W.xlsx")
    wb.save("C:/Users/hasegawa-t/Documents/シフト/作業工程_1W.xlsx")


if __name__ == '__main__':
    main()
