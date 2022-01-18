# -*- coding: utf-8 -*-
"""
Created on Sun Jan 16 15:36:16 2022

@author: Grant Huang
"""

import csv
import logging
from openpyxl import Workbook, load_workbook
import datetime
import re

months = ["january", "february", "march", "april", "may", "june", "july", 
          "august", "september", "october", "november", "december"]
month = ""
year = 0
wb = Workbook()

#letting user type in file names until they successfully open xl file
while True:
    try:
        #f = input("Input filename (.xlsx, .xlsm, .xltx, or .xltm file required): \n")
        f = "expedia_report_monthly_january_2018.xlsx"
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
            
        year = int(re.findall(r"\d{4}", f)[0])
        print(year)
        
        #found month in filename
        if month != "" and year != 0:
            break
        else:
            print("Can't find month and/or year in filename")

summaryRolling = wb["Summary Rolling MoM"]
monthSummary = ()
print(year)

#finding row containing data for month contained in filename
for row in summaryRolling.iter_rows(min_row = 1, max_row = 14, max_col = 6,
                                    values_only = True):
    #verifying leftmost column has a date
    if isinstance(row[0], datetime.date): 
        date = row[0]
        #print(date)
        #print(type(date))
        
        #found row corresponding to month in filename
        if date.strftime("%B").lower() == month:
            monthSummary = row
            break
    
print(monthSummary)
keys = ["date", "callsOffered", "abandon30s", "fcr", "dsat", "csat"]
monthData = {keys[i]: monthSummary[i] for i in range(len(keys))}
print(monthData)
print(type(monthData["abandon30s"]))

#logging all desired data from Summary Rolling MoM sheet
logging.basicConfig(filename="log.log", level=logging.DEBUG,
                    format="[%(levelname)s] %(asctime)s - %(message)s")
date = monthData["date"]
logging.info("%s %s Report", date.strftime("%B"), str(date.year))
logging.info("Calls Offered: %s", monthData["callsOffered"])
logging.info("Abandon after 30s: %.2f%%", monthData["abandon30s"]*100)
logging.info("FCR: %.2f%%", monthData["fcr"]*100)
logging.info("DSAT: %.2f%%", monthData["dsat"]*100)
logging.info("CSAT: %.2f%%", monthData["csat"]*100)

vocRolling = wb["VOC Rolling MoM"]
print(vocRolling)

#for row in vocRolling.iter_rows(min_row = 1, max_row = 14, max_col = 6,
#                                    values_only = True):
#    print(row)
    
colHeaders = vocRolling[1]
#for header in colHeaders:
#    print(header.value) 

col = 0

#searching for column with correct month and year
for n in range(len(colHeaders)):
    #found date
    if isinstance(colHeaders[n].value, datetime.date):
        headerDate = colHeaders[n].value
        headerMonth = headerDate.strftime("%B").lower()
    
        #found date and year in filename
        if headerMonth == month and headerDate.year == year:
            col = n
            break
        
rowHeaders = ["nps", "base size", "promoters", "passives", "dectractors", 
              "nps percent", "aarp total", "sat w/ agent percent", "aarp total",
              "dsat w/ agent percent", "aarp total"]
colData = [row[col] for row in vocRolling.iter_rows(values_only = True) 
           if row[0] is not None]
        
print(rowHeaders)
print(colData)


wb.close()