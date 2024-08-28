
import os
from PIL import Image
import pytesseract
import pandas as pd
import re

def image_to_excel(image_path, output_excel='data_table.xlsx'):
    # Step 1: Extract text from image using OCR
    image = Image.open(image_path)
    text = pytesseract.image_to_string(image)

    # Step 2: Parse the extracted text
    lines = text.splitlines()
    rows = []
    for line in lines:
        # Use regex to split the line, preserving all decimal places
        row = re.findall(r'-?\d+\.\d+|-?\d+|\S+', line)
        if row:
            rows.append(row)

    # Step 3: Create a DataFrame
    df = pd.DataFrame(rows)

    # Step 4: Save to Excel without changing data types or rounding
    df.to_excel(output_excel, index=False, float_format='%.15f')
    print(f"Excel file saved as {output_excel}")

# Example usage
image_path = r"C:\Users\ZANG\OneDrive\Desktop\DataExtraction\data_table.jpg"
output_excel = "data_table.xlsx"
image_to_excel(image_path, output_excel)
