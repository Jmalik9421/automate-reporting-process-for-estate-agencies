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
# ---------- DEFINING FUNCTIONS ----------- # 



# -------------------------- EDITING IMAGES -------------------------- #









