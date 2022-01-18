# -*- coding: utf-8 -*-
"""
Created on Sun Jan 16 15:36:16 2022

@author: Grant Huang
"""

import csv
from logging import Logger
from openpyxl import Workbook, load_workbook
import datetime

months = ["january", "february", "march", "april", "may", "june", "july", 
          "august", "september", "october", "november", "december"]
month = ""
wb = Workbook()

#letting user type in file names until they successfully open xl file
while True:
    try:
        f = input("Input filename (.xlsx, .xlsm, .xltx, or .xltm file required): \n")
        #f = "expedia_report_monthly_january_2018.xlsx"
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

summaryRolling = wb["Summary Rolling MoM"]
monthSummary = ()

#finding row containing data for month contained in filename
for row in summaryRolling.iter_rows(min_row = 1, max_row = 14, max_col = 6,
                                    values_only = True):
    #verifying leftmost column has a date
    if isinstance(row[0], datetime.date): 
        date = row[0]
        print(date)
        print(type(date))
        
        #found row corresponding to month in filename
        if date.strftime("%B").lower() == month:
            monthSummary = row
            break
    
print(monthSummary)
wb.close()