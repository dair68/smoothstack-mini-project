# -*- coding: utf-8 -*-
"""
Created on Sun Jan 16 15:36:16 2022

@author: Grant Huang
"""

import csv
from logging import Logger
from openpyxl import Workbook, load_workbook
import sys

months = ["january", "february", "march", "april", "may", "june", "july", 
          "august", "september", "october", "november", "december"]
month = ""
wb = Workbook()

#letting user type in file names until they successfully open xl file
while True:
    try:
        f = input("Input filename (.xlsx, .xlsm, .xltx, or .xltm file required): \n")
        #f = "expedia_report_monthly_janusary_2018.xlsx"
        #f = "expedia_report_monthly_march_2018.xlsx"
        wb = load_workbook(filename=f, read_only = True)
    except FileNotFoundError:
        print("Error: File not found")
    except Exception:
        print("Error: can't open file")
    else:
        #figuring out which month contained in filename
        for m in months:
            #found month contained within filename
            if f.find(m.lower()) != -1:
                month = m.lower()
                break
        
        #found month in filename
        if month != "":
            break
        else:
            print("Can't find month in filename")

ws = wb["Summary Rolling MoM"]
print(ws)
        
    
#found no month in filename
if month == "":
    print("Error: month not found in file name")
    wb.close()
    sys.exit()
    
print(m)    
wb.close()