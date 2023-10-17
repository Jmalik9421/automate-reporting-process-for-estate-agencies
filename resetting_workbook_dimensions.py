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
# ------- PRIMING GRID COORDINATES -------- # 



# ------- PRIMING GRID DIMENSIONS --------- # 
width_corrective_factor = (1 / 7)
height_corrective_factor = (100 / 133)

# --------- WIDTH --------- #
width_primed = pyip.inputNum(
    prompt = "Enter the width you want for each cell to be primed to (default: 64): ",
    greaterThan = 0
)

# VALIDATING INPUT #
while type(width_primed) != int:
    print(f"'{width_primed}' is not an integer.")
    width_primed = pyip.inputNum(
        prompt = "Enter the width you want for each cell to be primed to (default: 64): ",
        greaterThan = 0
    )
# VALIDATING INPUT #

width_primed *= width_corrective_factor
# --------- WIDTH --------- #



# --------- HEIGHT -------- #
height_primed = pyip.inputNum(
    prompt = "Enter the height you want for each cell to be primed to (default: 20): ",
    greaterThan = 0
)

# VALIDATING INPUT #
while type(height_primed) != int:
    print(f"'{height_primed}' is not an integer.")
    height_primed = pyip.inputNum(
        prompt = "Enter the height you want for each cell to be primed to (default: 20): ",
        greaterThan = 0
    )
# VALIDATING INPUT #

height_primed *= height_corrective_factor
# --------- HEIGHT -------- #

# ------- PRIMING GRID DIMENSIONS --------- # 








