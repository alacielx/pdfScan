## UPDATED 10/05/23 ##
currentVersion = 'v1.53'

import os
import re
import glob
import time
import pytesseract
from pytesseract import Output
import numpy as np
import subprocess
import cv2
import tkinter as tk
from tkinter import messagebox
import sys
import fitz  # PyMuPDF
from install_packages import download_zip
from functions import *


def getPdfPage(pdf_path, page_number):
    pdf_document = fitz.open(pdf_path)
    
    page = pdf_document.load_page(page_number)
    pixmap = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Increase resolution
    
    # Get image data as bytes
    image_data = pixmap.samples

    # Convert bytes to a NumPy array
    image_array = np.frombuffer(image_data, dtype=np.uint8).reshape(pixmap.height, pixmap.width, 3)

    pdf_document.close()
    
    return image_array

def runConvertFromPath(pdf):
    try:
        poppler_path = os.path.join(poppler_dir, 'pdftoppm')
        args = [
            poppler_path,
            '-f', str(1),
            '-l', str(1),
            '-jpeg',
            # '-r', str(600),  # Set the output resolution (default is 72 DPI)
            pdf,
        ]

        # Run the subprocess and capture the output
        process = subprocess.Popen(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=subprocess.CREATE_NO_WINDOW)
        output, error = process.communicate()

        # Convert the binary image data to a NumPy array
        image_data = np.frombuffer(output, np.uint8)
        image = cv2.imdecode(image_data, cv2.IMREAD_COLOR)

        return image

    except subprocess.CalledProcessError as e:
        print(f"Error occurred: {e}")
    except FileNotFoundError as e:
        print(f"Poppler not found: {e}")

def findText(text, data):
    
    # Get the bounding box for the given text
    data_words = data['text']
    text=text.split(' ')
    start_index = -1
    
    foundWords = False

    for i in range(len(data_words) - len(text) + 1):
        wordsSet = data_words[i:i+len(text)]
        for i2 in range(len(wordsSet)):
            if text == wordsSet:
                start_index = i
                end_index = i+len(text)-1
                foundWords = True
                break
        if foundWords:
            break
            
    if start_index == -1:
        return None

    left, bottom, right, top = None, None, None, None
    
    for i in range(start_index,end_index+1):
        left = data['left'][i] if left is None else min(data['left'][i],left)
        top = data['top'][i] if top is None else min(data['top'][i],top)
        right = data['left'][i] + data['width'][i] if right is None else max(data['left'][i] + data['width'][i],right)
        bottom = data['top'][i] + data['height'][i] if bottom is None else max(data['top'][i] + data['height'][i],bottom)
    
    width = right - left
    height = bottom - top

    if left is None or bottom is None or right is None or top is None:
        raise ValueError(f"Failed to find the bounding box for '{text}' text.")
    
    return left, top, width, height

def expandAndCrop(image, box, width, height, wpadding, hpadding):
    
    try:
        (box_left, box_top, box_width, box_height) = (box[0],box[1],box[2],box[3])
    except:
        # Handle the exception if 'box' doesn't have exactly four elements
        print("Error: 'box' should contain exactly four elements.")
        return None

    new_width = box_width + int(box_width*(width/100))
    new_height = box_height + int(box_height*(height/100))
    wpadding = wpadding/100
    hpadding = hpadding/100

    box_left = box_left - (int(new_width*wpadding/2))
    new_width = new_width + int(new_width*wpadding)
    box_top = box_top - (int(new_height*hpadding/2))
    new_height = new_height + int(new_height*hpadding)
    cropped_image = image[box_top:box_top + new_height, box_left:box_left + new_width]

    return cropped_image

def preprocessImage(image, threshold):
    # Apply linear contrast adjustment using the formula: output_image = alpha * input_image + beta
    _, image = cv2.threshold(image, threshold, 255, cv2.THRESH_BINARY)

    image = cv2.medianBlur(image, 3)

    return image

def adaptiveThresholding(image, output_filename):
    adaptive_image = cv2.adaptiveThreshold(image, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 11, 5)
    cv2.imwrite(output_filename, adaptive_image)

def adjustContrast(image, contrast_factor):
    # Ensure that the contrast factor is within the valid range (typically between 0 and 2)
    contrast_factor = max(0, contrast_factor)
    contrast_factor = min(2, contrast_factor)

    # Convert the image to floating-point to perform the contrast adjustment
    image = image.astype(np.float32)

    # Apply the contrast adjustment
    adjusted_image = (image - 128) * contrast_factor + 128

    # Clip the pixel values to ensure they remain within the valid range [0, 255]
    adjusted_image = np.clip(adjusted_image, 0, 255)

    # Convert the image back to unsigned 8-bit integer (grayscale) format
    adjusted_image = adjusted_image.astype(np.uint8)

    return adjusted_image

#####################################################################################################

updateExecutable(currentVersion, "pdfScan")

root = tk.Tk()
root.withdraw()

#Check if config file exists
# config_file_name = 'pdfScanConfig.ini'
# if not os.path.exists(config_file_name):
#     create_config_file()

#Check if config file exists and has all options
config_file_name = 'pdfScanConfig.ini'
configProps = {"pdf_folder" : "", "working_date" : "", "installation_date" : "", "initials" : "", "add_so_number" : "True"}

checkConfig(config_file_name, configProps)
configProps = readConfig(config_file_name)

#Check installation date
today = time.strftime("%m.%d")

if not configProps['working_date'] == today:
    while True:
        configProps['installation_date'] = askInput("Enter Installation Date (ie. XX.XX):","PDF Scanning " + currentVersion)
        installPattern = r"^\d{2}\.\d{2}$"
        if re.search(installPattern,configProps['installation_date']):
            configProps['working_date'] = today
            break
        messagebox.showinfo("PDF Scanning " + currentVersion, "Please enter a valid installation date")

if not configProps['pdf_folder']:
    configProps['pdf_folder'] = askFolderDirectory()

if not configProps['initials']:
    configProps['initials'] = askInput("Enter Initials:")

pdf_folder = configProps["pdf_folder"]
working_date = configProps["working_date"]
installation_date = configProps["installation_date"]
initials = configProps["initials"]
add_so_number = configProps["add_so_number"]

updateConfig(config_file_name, configProps)

# #Read config file
# pdf_folder, working_date, installation_date, initials, add_so_number = read_config_file()



#Change to test mode if in "Test" folder
# current_directory = os.getcwd()
# folders = current_directory.split("\\")
# if folders[len(folders)-1] == "Test":
#     pdf_folder = fr"{current_directory}\pdfScan"

poppler_dir = r"C:\poppler\poppler-23.07.0\Library\bin"

duplicate_count = 1
not_renamed_count = 0

if not os.path.exists("C:\tesseract\Tesseract\tesseract.exe"):
    tesseract = {'tesseract': "https://github.com/alacielx/Tesseract-Portable/raw/main/Tesseract.zip"}
    download_zip(tesseract)

pytesseract.pytesseract.tesseract_cmd = fr"C:\tesseract\Tesseract\tesseract.exe"

for pdf in glob.glob(os.path.join(pdf_folder, "*.pdf")):

    if os.path.basename(pdf).startswith("~$"):
        continue  # Skip temporary files
    
    # Read pdf as numpy array
    try:
        image = getPdfPage(pdf,0)
    except:
        continue
    try:
        image = image[1:int(image.shape[0]/2),int(image.shape[1]/5*2):image.shape[1]]
        image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    except:
        continue

    preprocessed_image = cv2.GaussianBlur(image, (3, 3), 0)
    preprocess_image_threashold = 110

    so_text = ""
    so_number = ""
    address_text = ""
    address = ""

    so_number_list = [[],[]]
    address_list = [[],[]]

    address_pattern = r"\d+[-\d]*\s+\w+(?:-\w+)?(?:\s+\w*)?"

    for i in np.arange(1.0,2.1,0.25):
        
        data = pytesseract.image_to_data(preprocessed_image, output_type=Output.DICT)
        
        # Get Sales Order Number
        so_box = findText("SALES ORDER", data)
        if so_box is not None and so_number == "":
            so_image = expandAndCrop(preprocessed_image, so_box, 200, 0, 5, 25)
            so_data = pytesseract.image_to_data(so_image, output_type=Output.DICT, config=f"--psm 13")
            
            for index, element in enumerate(so_data['text']):
                match = re.search(r"^\d{8}$", element)
                confidence = so_data['conf'][index]
                
                if match and len(element) == 8:
                    if confidence > 80:
                        so_number = element
                        break
                    else:
                        so_number_list[0].append(element)
                        so_number_list[1].append(confidence)
                else:
                    continue


        # Get Address
        address_box = findText("to:", data)
        if address_box is None:
            address_box = findText("to;", data)

        if address_box is not None and address == "":
            address_image = expandAndCrop(preprocessed_image, address_box, 445, 750, 10, 20)
            address_image = expandAndCrop(preprocessed_image, address_box, 1300, 1000, 50, 20)
            
            address_data = pytesseract.image_to_data(address_image, output_type=Output.DICT, config=f"--psm 6")

            address_text = " ".join(address_data['text'])

            # Find the address pattern in the string
            address_match = re.search(address_pattern, address_text)
            if not address_text[address_match.start() - 1] == " ":
                address_match = None

            # Get address if conf >90, otherwise add to list to pick highest confidence later
            if address_match:
                address_text = address_match.group().split(" ")
            

                start_index = address_data['text'].index(address_text[0])
                confidence = address_data['conf'][start_index:start_index + len(address_text)]
                confidence = sum(confidence) / len(confidence)
                
                if confidence > 90:
                    address = address_match.group()
                    address = sanitizeName(address)
                else:
                    address_list[0].append(address_match.group())
                    address_list[1].append(confidence)

        if not so_number == "" and not address == "":
            break
        
        preprocessed_image = adjustContrast(image,i)

    if so_number == "" and so_number_list[0]:
        max_index = so_number_list[1].index(max(so_number_list[1]))
        so_number = so_number_list[0][max_index]
    
    if address == "" and address_list[0]:
        max_index = address_list[1].index(max(address_list[1]))
        address = address_list[0][max_index]

    
    new_file_name = f"{address}-{installation_date}-{initials}"

    if add_so_number == 'True':
        new_file_name = f"{new_file_name}-{so_number}"

    new_file_path = os.path.join(pdf_folder,f"{new_file_name}.pdf")

    while os.path.exists(new_file_path):
        new_file_path = os.path.join(pdf_folder,f"{new_file_name}_{duplicate_count}.pdf")
        duplicate_count += 1
    
    original_pdf_name = os.path.splitext(os.path.basename(pdf))[0]

    if not address == "" and not so_number == "":
        if not original_pdf_name == new_file_name:
            try:
                os.rename(pdf, new_file_path)
            except:
                not_renamed_count += 1
                continue
    else:
        not_renamed_count += 1

if not_renamed_count == 0:
    message = "Done"
elif not_renamed_count == 1:
    message = f"{not_renamed_count} file was unable to be renamed :( \nPlease check."
else:
    message = f"{not_renamed_count} files were unable to be renamed :( \nPlease check."

messagebox.showinfo("PDF Scanning " + currentVersion, message)