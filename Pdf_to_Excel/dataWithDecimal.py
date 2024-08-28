import os
from pdf2image import convert_from_path
import pytesseract
import pandas as pd
import re

# ... (previous functions remain unchanged)

# Step 3: Parse Text and Save to Excel
def text_to_excel(text, output_excel='output.xlsx'):
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