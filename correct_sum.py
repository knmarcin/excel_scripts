# It's common in excel that when you are moving rows, =SUM formulas aren't necessarily corrected. 
# Especially when you are moving row to a first or last position.
# This script finds all =SUM formulas and it corrects them in curent sheet. 

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
checkIfSum =''

wb = openpyxl.load_workbook(file_direction)
sheet = wb[JakiSheet]

wbb = excel.Workbooks.Open(file_direction)
wss = wbb.Worksheets(JakiSheet)

while iteracja>30:
    if sheet[('I' + str(iteracja))].value!=None:
        checkIfSum = str(sheet[('I' + str(iteracja))].value)[:4]
        if checkIfSum == '=SUM':
            sumLocation = iteracja
            PierwszaWartosc = iteracja-1

            while (sheet[('I' + str(iteracja))].value != None):
                iteracja = iteracja - 1
                
            LastValue = iteracja + 1    
            print('last value: ',iteracja+1)
            print('Pierwsza wartosc: ',PierwszaWartosc)
            print('Sum location: ',sumLocation)
            
            RANGE = 'I'+str(LastValue)+':I'+str(PierwszaWartosc)
            print(RANGE)
            wss.Cells(sumLocation,9).FORMULA = '=SUM('+str(RANGE)+')'
        else:
            itearcja = iteracja - 1
    
    iteracja = iteracja - 1      
