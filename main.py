# -*- coding: utf-8 -*-
"""
Created on Sun Jan 16 15:36:16 2022

@author: Grant Huang
"""

import csv
from logging import Logger
from openpyxl import Workbook, load_workbook
import sys


wb = Workbook()

#letting user type in file names until they successfully open xl file
while True:
    try:
        f = input("Input filename: \n")
        #f = "expedia_report_monthly_january_2018.xlsx"
        #f = "expedia_report_monthly_march_2018.xlsx"
        wb = load_workbook(filename=f, read_only = True)
        break
    except FileNotFoundError:
        print("Error: File does not exist")
    except Exception:
        print("Error: can't open file")

ws = wb["Summary Rolling MoM"]
print(ws)
        
months = ["january", "february", "march", "april", "may", "june", "july", 
          "august", "september", "october", "november", "december"]
month = ""

#figuring out which month contained in filename
for m in months:
    #found month contained within filename
    if f.find(m) != -1:
        month = m
        break
    
#found no month in filename
if month == "":
    print("Error: month not found in file name")
    wb.close()
    sys.exit()
    
print(m)    
wb.close()