import fitz  # PyMuPDF
import pdfreader
import pytesseract
import pandas as pd
import camelot

def pdf_to_images(pdf_path):
    """Converts PDF pages to images."""
    images = pdfreader.convert_from_path(pdf_path)
    return images

def extract_text_from_image(image):
    """Uses OCR to extract text from an image."""
    text = pytesseract.image_to_string(image)
    return text

def convert_pdf_to_excel(pdf_path, output_excel_path):
    """Converts a PDF file containing tables directly to Excel using Camelot."""
    # Using Camelot to read PDF tables
    tables = camelot.read_pdf(pdf_path, pages='all', flavor='stream')
    
    # If tables are detected, export them to Excel
    if tables:
        tables.export(output_excel_path, f='excel')
        print(f"Successfully extracted tables to {output_excel_path}")
    else:
        print("No tables found in PDF, attempting image-based OCR.")
        
        # Convert PDF to images if no tables found
        images = pdf_to_images(pdf_path)
        
        # Extract text from images and process to Excel
        data = []
        for img in images:
            text = extract_text_from_image(img)
            # Process text to structured data format (implementation specific to the text structure)
            # Example: data.append(process_text_to_data(text))
        
        # Convert extracted text data to DataFrame and export to Excel
        df = pd.DataFrame(data)
        df.to_excel(output_excel_path, index=False)
        print(f"Successfully extracted OCR text data to {output_excel_path}")

if __name__ == "__main__":
    pdf_path = r"C:\Users\ZANG\OneDrive\Desktop\DataExtraction\extracted_pages.pdf"
    output_excel_path = "data_extracted.xlsx"
    convert_pdf_to_excel(pdf_path, output_excel_path)
