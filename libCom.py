#!/usr/bin/env python
# coding: utf-8

# In[1]:


#リストを1行ずつprintする
def print_list(list):
    for i in range(len(list)):
        print(list[i])


# In[2]:


#シートのセル幅を自動で整える
def auto_sheet_width(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column

        for cell in col:
            if get_east_asian_width_count(str(cell.value)) > max_length:
                max_length = get_east_asian_width_count(str(cell.value))

        adjusted_width = (max_length + 2) * 1.0
        ws.column_dimensions[col[0].column_letter].width = adjusted_width
        
        
#半角を1文字、全角2を文字としてカウントする
import unicodedata
def get_east_asian_width_count(text):
    count = 0
    for c in text:
        if unicodedata.east_asian_width(c) in 'FWA':
            count += 2
        else:
            count += 1
    return count

