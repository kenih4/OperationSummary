#!/usr/bin/python
# -*- coding: utf-8 -*-

import libCom
import schedule
import BlFaultSummary
import datetime
import sys
import os


bl = input(">BLを選択してください。(bl1,bl2,bl3)>>>")
if bl == "bl1":
    acc = 1
else:
    acc = 0
    
fault_list = BlFaultSummary.get_fault_list(acc)
list = BlFaultSummary.get_unit_list(bl,fault_list)
libCom.print_list(list)
BlFaultSummary.output_log_txt(list)