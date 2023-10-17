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



# ----------- PRIMING INSERTION ----------- # 
alignment_input = pyip.inputChoice(
    prompt = "Would you like for the images to be inserted vertically or horizontally?: ",
    choices = ["vertically", "v", "horizontally", "v"]
)

column_input = pyip.inputChoice(
    prompt = "Please enter the column you would like to start inserting the images on [A-Z, AA-ZZ]: ",
    choices = columns
)

row_input = pyip.inputNum(
    prompt = "Please enter the row you would like to start inserting the images on [1-702]: ",
    greaterThan = 0
)

while type(row_input) != int:
    print(f"'{row_input}' is not an integer.")
    row_input = pyip.inputNum(
        prompt = "Please enter desired width: ",
        greaterThan = 0
    )
# ----------- PRIMING INSERTION ----------- # 



# ------------- OPENING SHEET ------------- # 
workbook = openpyxl.load_workbook(file_workbook_abs_path)
worksheet_active = workbook.active
# ------------- OPENING SHEET ------------- # 



# ------------- PRIMING IMAGES ------------ # 
greatest_width_list = []
for i in list(copies_images_dict.values()):
    greatest_width_list.append(i[1])

greatest_height_list = []
for i in list(copies_images_dict.values()):
    greatest_height_list.append(i[2])

def greatest_value(lst):
    lst.sort()
    return lst[-1]

width_corrective_factor = (1 / 7)
height_corrective_factor = (100 / 133)

greatest_width = greatest_value(greatest_width_list) * width_corrective_factor      # corrective factor as excel increases dimensions substantially. Testing confirmed not due to code. 
greatest_height = greatest_value(greatest_height_list) * height_corrective_factor   # corrective factor as excel increases dimensions substantially. Testing confirmed not due to code. 
# ------------- PRIMING IMAGES ------------ # 



# ------------ INSERTING IMAGES ----------- # 
def insert_vertically():
    for i in range(len(copies_images_dict)):
        # EDITING GRID DIMENSIONS #
        worksheet_active.column_dimensions[column_input].width = greatest_width
        worksheet_active.row_dimensions[row_input + 1].height = greatest_height
        # EDITING GRID DIMENSIONS #

        current_image_path = dir_copies_images_abs_path + "\\" + f"{copies_images_dict[i][0]}"
        current_image = Image_openpyxl(current_image_path)
        worksheet_active.add_image(
            current_image,
            f"{column_input}{row_input + i}"                                        # columns stay constant
        )

def insert_horizontally():
    for i in range(len(copies_images_dict)):
        # EDITING GRID DIMENSIONS #
        worksheet_active.column_dimensions[columns[columns.index(column_input) + i]].width = greatest_width
        worksheet_active.row_dimensions[row_input].height = greatest_height
        # EDITING GRID DIMENSIONS #

        current_image_path = dir_copies_images_abs_path + "\\" + f"{copies_images_dict[i][0]}"
        current_image = Image_openpyxl(current_image_path)
        worksheet_active.add_image(
            current_image,
            f"{columns[columns.index(column_input) + i]}{row_input + i}"            # rows stay constant                                # columns stay constant
        )

if alignment_input == "vertically" or alignment_input == "v":
    insert_vertically

elif alignment_input == "horizontally" or alignment_input == "h":
    insert_horizontally

workbook.save(file_workbook_abs_path)
# ------------ INSERTING IMAGES ----------- # 
# ------------------- INSERTING IMAGES TO WORKBOOK ------------------- #









