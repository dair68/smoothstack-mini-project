# -*- coding: utf-8 -*-
"""
Created on Sun Jan 16 15:36:16 2022

@author: Grant Huang
"""

import csv
from logging import Logger
from openpyxl import load_workbook

#january filename: expedia_report_monthly_january_2018.xlsx
f = input("Input filename: \n")
print(f)

wb = load_workbook(filename=f, read_only = True)
ws = wb["Summary Rolling MoM"]
print(ws)

for row in ws.rows:
    for cell in row:
        print(cell.value)