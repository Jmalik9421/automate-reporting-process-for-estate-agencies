# ------------------------ IMPORTING MODULES ------------------------ #
import os
import shutil
import openpyxl
import pyinputplus as pyip
from PIL import Image as Image_PIL
from openpyxl.drawing.image import Image as Image_openpyxl
# ------------------------ IMPORTING MODULES ------------------------ #



# ----------------------- PRIMING DIRECTORIES ----------------------- #
dir_og_images_abs_path = os.path.abspath(".\\Images\\Originals")
dir_copies_images_abs_path = os.path.abspath(".\\Images\\Copies")

file_workbook_abs_path = os.path.abspath(".\\Spreadsheets\\workbook.xlsx")

og_images_list = os.listdir(dir_og_images_abs_path)
copies_images_list = os.listdir(dir_copies_images_abs_path)
# ----------------------- PRIMING DIRECTORIES ----------------------- #



# -------------------------- EDITING IMAGES -------------------------- #
# ---------- DEFINING FUNCTIONS ----------- # 
def reset():
    for i in copies_images_list:
        copies_current_image = dir_copies_images_abs_path + f"\\{i}"
        os.remove(copies_current_image)
    
    for i in og_images_list:
        og_current_image = dir_og_images_abs_path + f"\\{i}"
        shutil.copy(
            og_current_image, 
            dir_copies_images_abs_path
        )
    
    return (
        os.listdir(dir_og_images_abs_path),
        os.listdir(dir_copies_images_abs_path)
    )

def resize():
    width_input = pyip.inputNum(
        prompt = "Please enter desired width: ",
        greaterThan = 0
    )
    while type(width_input) != int:
        print(f"'{width_input}' is not an integer.")
        width_input = pyip.inputNum(
            prompt = "Please enter desired width: ",
            greaterThan = 0
        )
    
    height_input = pyip.inputNum(
        prompt = "Please enter desired height: ",
        greaterThan = 0
    )
    while type(height_input) != int:
        print(f"'{height_input}' is not an integer.")
        height_input = pyip.inputNum(
            prompt = "Please enter desired height: ",
            greaterThan = 0
        )

    for item, image in enumerate(copies_images_list):
        copies_current_image = dir_copies_images_abs_path + f"\\{image}" 
        current_image = Image_PIL.open(copies_current_image)            

        if width_input > current_image.width or height_input > current_image.height:
            print(f"'{image}' is being upsacled. This will cause it to be stretched. Image quality will suffer.")

        new_image = current_image.resize((width_input, height_input)) 
        new_image.save(
            dir_copies_images_abs_path
            + "\\"
            + f"{width_input}x{height_input}_image{item}.jpg"
        )
        os.remove(copies_current_image)
    
    return(
        os.listdir(dir_og_images_abs_path),
        os.listdir(dir_copies_images_abs_path)
    )
# ---------- DEFINING FUNCTIONS ----------- # 



# ------------ RESET REQUIRED ------------- # 
if og_images_list != copies_images_list:
    print("Images have been edited.")
    reset_input = pyip.inputYesNo("Would you like to reset?: ")

    if reset_input == "yes":
        (og_images_list, copies_images_list) = reset()                  # re-assign variables to unpacked returned tuple which contains new named images in respective directories.
        (og_images_list, copies_images_list) = resize()
    else:
        resize_input = pyip.inputYesNo("Would you like to resize the images?: ")
        if resize_input == "yes":
            (og_images_list, copies_images_list) = resize()
# ------------ RESET REQUIRED ------------- # 


# ---------- RESET NOT REQUIRED ----------- # 
else:
    resize_input = pyip.inputYesNo("Would you like to resize the images?: ")
    if resize_input == "yes":
        (og_images_list, copies_images_list) = resize()
# ---------- RESET NOT REQUIRED ----------- # 



copies_images_dict = {}
for index,name in enumerate(copies_images_list):
    copies_current_image = dir_copies_images_abs_path + f"\\{name}"
    current_image = Image_PIL.open(copies_current_image)
    width = current_image.width
    height = current_image.height
    copies_images_dict.setdefault(index, [name, width, height])
# -------------------------- EDITING IMAGES -------------------------- #



# ------------------- INSERTING IMAGES TO WORKBOOK ------------------- #
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

# ------------------- INSERTING IMAGES TO WORKBOOK ------------------- #









