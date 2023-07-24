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
from tkinter import simpledialog
from tkinter import filedialog
from tkinter import messagebox
import sys
import configparser
from PIL import Image
import io

test = False

# Function to sanitize file names
def sanitize_name(file_name):
    
    file_name = file_name.split("\n")[0]

    invalid_chars = ["<", ">", ":", '"', "/", "\\", "|", "?", "*"]
    
    # Replace invalid characters and remove at the end of the file name
    for invalid_char in invalid_chars:
        if file_name.endswith(invalid_char) or file_name.endswith(" "):
            file_name = file_name[:-len(invalid_char)]
        else:
            file_name = file_name.replace(invalid_char, "_")
    
    return file_name

def run_convert_from_path(pdf):
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

def image_to_data(np_image):
    try:
        tesseract_cmd = fr"C:\tesseract\Tesseract\tesseract.exe"
        # Convert the NumPy image to a PIL Image
        pil_image = Image.fromarray(np_image)

        # Save the PIL Image to a temporary file (in memory)
        temp_image = io.BytesIO()
        pil_image.save(temp_image, format="PNG")
        temp_image.seek(0)

        # Read the image data as bytes
        image_data_bytes = temp_image.read()

        # Construct the command to run Tesseract with desired options
        command = [
            tesseract_cmd,
            "stdin",  # Use stdin to read the image from a file-like object
            "stdout",
            "output_type=Output.DICT",
            "--dpi",
            "300",
            "--oem",
            "3",
            "--psm",
            "6",
            "-l",
            "eng",
            "hocr"
        ]

        # Run Tesseract using subprocess and capture the output
        process = subprocess.Popen(command, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=subprocess.CREATE_NO_WINDOW)
        extracted_data, _ = process.communicate(input=image_data_bytes)

        # Decode the extracted_data to string
        extracted_data = extracted_data.decode()

        return extracted_data

    except subprocess.CalledProcessError as e:
        print(f"Error occurred: {e}")
    except FileNotFoundError as e:
        print(f"Tesseract not found: {e}")

def ask_folder_directory():
    root = tk.Tk()
    root.withdraw()
    folder_selected = ""

    while not folder_selected:
        folder_selected = filedialog.askdirectory(title="Select PDF folder")
        return os.path.normpath(folder_selected)
    else:
        sys.exit()


def create_config_file():

    pdf_folder = ask_folder_directory()
    initials = ask_initials()
    config = configparser.ConfigParser()
    config['DEFAULT'] = {
        'pdf_folder': pdf_folder,
        'working_date': "",
        'installation_date': "",
        'initials': initials,
        'add_so_number': True
    }

    with open(config_file_name, 'w') as configfile:
        config.write(configfile)
    

def read_config_file():
    config = configparser.ConfigParser()
    config.read(config_file_name)

    pdf_folder = config['DEFAULT']['pdf_folder']
    working_date = config['DEFAULT']['working_date']
    installation_date = config['DEFAULT']['installation_date']
    initials = config['DEFAULT']['initials']
    add_so_number = config['DEFAULT']['add_so_number']

    return pdf_folder, working_date, installation_date, initials, add_so_number

def ask_installation_date():
    config = configparser.ConfigParser()
    config.read(config_file_name)
    
    root = tk.Tk()
    root.withdraw()

    windowTitle = " "
    askInstallationMessage = "Enter Installation Date (ie. XX.XX):"

    new_installation_date = simpledialog.askstring(windowTitle, askInstallationMessage)

    if new_installation_date is None:
        sys.exit()

    config.set("DEFAULT","installation_date",new_installation_date)
    config.set("DEFAULT","working_date",today)

    with open(config_file_name, "w") as configfile:
        config.write(configfile)

def ask_initials():
    root = tk.Tk()
    root.withdraw()

    windowTitle = " "
    askInstallationMessage = "Enter Initials:"

    initials = simpledialog.askstring(windowTitle, askInstallationMessage)

    if initials is None:
        sys.exit()

    return initials

def find_text(text, data):
    
    # Get the bounding box for the given text
    data_words = data['text']
    text=text.split(' ')
    start_index = -1

    for i in range(len(data_words) - len(text) + 1):
        if data_words[i:i+len(text)] == text:
            start_index = i
            end_index = i+len(text)-1
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
    # cropped_image = image[top:top+height, left:left+width]
    # cv2.imshow("img", cropped_image)
    # cv2.waitKey(0)
    return left, top, width, height

def find_text2(text, data):
    
    # Get the bounding box for the given text
    data_words = data['text']
    text=text.split(' ')
    word_indexes = []

    for i in text:
        try:
            word_indexes.append(data_words.index(i))
        except:
            return None
    
    if len(word_indexes) == 0:
        return None

    left, bottom, right, top = None, None, None, None
    
    for i in word_indexes:
        left = data['left'][i] if left is None else min(data['left'][i],left)
        top = data['top'][i] if top is None else min(data['top'][i],top)
        right = data['left'][i] + data['width'][i] if right is None else max(data['left'][i] + data['width'][i],right)
        bottom = data['top'][i] + data['height'][i] if bottom is None else max(data['top'][i] + data['height'][i],bottom)
    
    width = right - left
    height = bottom - top

    if left is None or bottom is None or right is None or top is None:
        raise ValueError(f"Failed to find the bounding box for '{text}' text.")
    # cropped_image = image[top:top+height, left:left+width]
    # cv2.imshow("img", cropped_image)
    # cv2.waitKey(0)
    return left, top, width, height

def expand_and_crop(image, box, width, height, wpadding, hpadding):
    
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
    # cv2.imshow("img", cropped_image)
    # cv2.waitKey(0)
    return cropped_image

def preprocess_image(image, threshold=110):
    image = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    # Apply linear contrast adjustment using the formula: output_image = alpha * input_image + beta
    # image = cv2.convertScaleAbs(image, alpha=1.7, beta=0)
    _, image = cv2.threshold(image, threshold, 255, cv2.THRESH_BINARY)

    return image

#####################################################################################################

#Check if config file exists
config_file_name = 'pdfScanConfig.ini'
if not os.path.exists(config_file_name):
    create_config_file()

#Read config file
pdf_folder, working_date, installation_date, initials, add_so_number = read_config_file()

#Check installation date
today = time.strftime("%m.%d")
if not working_date == today:
    ask_installation_date()
    pdf_folder, working_date, installation_date, initials, add_so_number = read_config_file()

#Change dir if in test mode
if test:
    current_folder = os.getcwd()
    pdf_folder = fr"{current_folder}\pdfScan"

poppler_dir = r"C:\poppler\poppler-23.07.0\Library\bin"

count = 1

pytesseract.pytesseract.tesseract_cmd = fr"C:\tesseract\Tesseract\tesseract.exe"

for pdf in glob.glob(os.path.join(pdf_folder, "*.pdf")):
    #run convert_from_path without console pdf to image
    image = run_convert_from_path(pdf)
    image = image[1:int(image.shape[0]/2),int(image.shape[1]/5*2):image.shape[1]]
    cv2.imwrite("img.jpg", image)
    image = preprocess_image(image)
    
    #MAKE THIS NOT SHOW CONSOLE
    data = pytesseract.image_to_data(image, output_type=Output.DICT)

    sales_order_box = find_text("SALES ORDER", data)
    if sales_order_box is not None:
        sales_order_image = expand_and_crop(image, sales_order_box, 200, 0, 5, 25)
        cv2.imwrite("img_sales_order.jpg", sales_order_image)
        sales_order_text = pytesseract.image_to_string(sales_order_image, config=f"--psm 13")
        sales_order_text = sales_order_text.replace('\n','')
    else:
        sales_order_text = ""

    address_box = find_text("to:", data)
    if address_box is not None:
        address_image = expand_and_crop(image, address_box, 445, 750, 10, 20)
        address_image = expand_and_crop(image, address_box, 1300, 1000, 50, 20)
        cv2.imwrite("img_address.jpg", address_image)
        address_text = pytesseract.image_to_string(address_image, config=f"--psm 6")
    else:
        address_text = ""
    
    sales_order_number = re.search(r'\d+', sales_order_text)
    
    if sales_order_number:
        sales_order_number = sales_order_number[0]
    else:
        sales_order_number = ""

    address_pattern = r"\d+[-\d]*\s+\w+(?:-\w+)?(?:\s+\w*)?"

    match = re.search(address_pattern, address_text)
    if match:
        address = match[0]
    else:
        address = ""

    address = sanitize_name(address) 
    
    new_file_name = f"{address}-{installation_date}-{initials}"

    if add_so_number == 'True':
        new_file_name = f"{new_file_name}-{sales_order_number}"

    new_file_path = os.path.join(pdf_folder,f"{new_file_name}.pdf")

    while os.path.exists(new_file_path):
        new_file_path = os.path.join(pdf_folder,f"{new_file_name}_{count}.pdf")
        count += 1
    
    if not address == "":
        os.rename(pdf, new_file_path)

messagebox.showinfo("PDF Scanning", "Done")