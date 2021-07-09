# Simple script counting and inserting number of documents for drivers.

import pandas as pd
import openpyxl
import xlrd
import numpy as np
import os
import py2win
import win32com.client
import warnings

warnings.filterwarnings("ignore")
excel = win32com.client.Dispatch("Excel.Application")
wb1 = pd.ExcelFile('\\\\tabo-srv1\\logist\\dokumenty.xls')
sheet = pd.read_excel(wb1, sheet_name="Arkusz1", usecols=["DATA"])
df = sheet
nazwa_arkusza = df['DATA'].iloc[0]

folder_name = nazwa_arkusza[-4:]
file_name = nazwa_arkusza[-7:]+'.xlsm'
file_direction = 'T:\\poziom 0\\optima_analizy_tabo upg\\logist\\SPECYFIKACJA\\'+ folder_name +'\\'+ file_name

JakiSheet = nazwa_arkusza
wb1.close()

PierwszaWartosc = ''
DrugaWartosc = ''

print(JakiSheet)

iteracja = 3500

wb = openpyxl.load_workbook(file_direction)
sheet = wb[JakiSheet]

wbb = excel.Workbooks.Open(file_direction)
wss = wbb.Worksheets(JakiSheet)

while iteracja!=1:
    if (sheet[('S' + str(iteracja))].value == None):
        iteracja = iteracja - 1
        counter = 1
    else:
        while((sheet[('S' + str(iteracja))].value != None) and (iteracja > 33)) :
                iteracja = iteracja - 1     
                counter = counter + 1
        iteracja = iteracja - 1
        wss.Cells(iteracja,1).Value = counter-1
        print(counter-1)
        
