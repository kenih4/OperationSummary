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
    print("正常終了しました")

    subprocess.Popen(['start', r"C:\Users\kenichi\Documents\OperationSummary\Modified\計画時間用マクロ.xlsm"], shell=True)

    print("schedule.計画時間ファイル:",schedule.計画時間ファイル)
    EXCEL = schedule.計画時間ファイル
    subprocess.Popen(['start', EXCEL], shell=True)    
else:
    print("異常終了しました")
time.sleep(60)