import os
from pdf2image import convert_from_path
import pytesseract
import pandas as pd

# Step 1: Convert PDF to Images
def pdf_to_images(pdf_path, output_folder='images', dpi=300):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    images = convert_from_path(pdf_path, dpi=dpi)
    image_paths = []
    for i, image in enumerate(images):
        image_path = os.path.join(output_folder, f'page_{i + 1}.png')
        image.save(image_path, 'PNG')
        image_paths.append(image_path)
    return image_paths

# Step 2: Extract Text from Images using OCR
def image_to_text(image_path):
    text = pytesseract.image_to_string(image_path)
    return text

# Step 3: Parse Text and Save to Excel
# def text_to_excel(text, output_excel='output.xlsx'):
#     # For simplicity, we'll split the text into rows and columns by new lines and spaces
#     rows = [line.split() for line in text.splitlines() if line.strip()]
    
#     # Creating a DataFrame
#     df = pd.DataFrame(rows)
    
#     # Saving to Excel
#     df.to_excel(output_excel, index=False)

import os
from pdf2image import convert_from_path
import pytesseract
import pandas as pd
import re

# ... (previous functions remain unchanged)

# Step 3: Parse Text and Save to Excel
def text_to_excel(text, output_excel='page291.xlsx'):
    # Split the text into lines
    lines = text.splitlines()
    
    # Parse each line, preserving decimal numbers
    rows = []
    for line in lines:
        # Use regex to split the line, keeping decimal numbers intact
        row = re.findall(r'\d+\.\d+|\S+', line)
        if row:
            rows.append(row)
    
    # Creating a DataFrame
    df = pd.DataFrame(rows)
    
    # Saving to Excel without changing data types
    df.to_excel(output_excel, index=False, float_format='%.15f')

# ... (rest of the code remains unchanged)

# Combined Function to Convert PDF to Excel
def pdf_to_excel(pdf_path, output_excel='page291.xlsx'):
    images = pdf_to_images(pdf_path)
    full_text = ""
    for image_path in images:
        text = image_to_text(image_path)
        full_text += text + "\n"
    
    text_to_excel(full_text, output_excel)
    print(f"Excel file saved as {output_excel}")

# Example Usage
pdf_path = r"C:\Users\ZANG\OneDrive\Desktop\DataExtraction\extracted_page291.pdf"
output_excel = 'page291.xlsx'
pdf_to_excel(pdf_path, output_excel)
