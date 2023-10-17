import os
import shutil
import openpyxl
import pyinputplus as pyip

# ----------------------- PRIMING DIRECTORIES ----------------------- #
file_workbook_abs = os.path.abspath(".\Spreadsheets\\workbook.xlsx")
# ----------------------- PRIMING DIRECTORIES ----------------------- #

# ----------------------- RESETTING WORKBOOK ------------------------ #

# ------------- OPENING SHEET ------------- # 
workbook = openpyxl.load_workbook(file_workbook_abs)
worksheet_active = workbook.active
# ------------- OPENING SHEET ------------- # 

# ------- PRIMING GRID COORDINATES -------- # 
# -- COLUMNS -- #
columns = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
columns1 = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
columns2 = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")

for letter1 in columns1:
    for letter2 in columns2:
        columns.append(letter1+letter2)                                 # 702 columns
# -- COLUMNS -- #

# --- ROWS ---- #
rows = []                                                               # 702 rows
for i in range(len(columns)):
    i += 1                                                              # 1 index rows as in workbook
    rows.append(i)
# --- ROWS ---- #







