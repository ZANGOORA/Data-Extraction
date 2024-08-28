import os
import openai
from PyPDF2 import PdfReader
import pandas as pd

# Set up OpenAI API key
openai.api_key = os.getenv("OPENAI_API_KEY")

def extract_text_from_pdf(pdf_path):
    """Extract text from PDF file."""
    reader = PdfReader(pdf_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text

def process_text_with_openai(text):
    """Process text using OpenAI API to extract structured data."""
    prompt = f"Extract structured data from the following text and format it as a list of dictionaries:\n\n{text}"
    
    response = openai.Completion.create(
        engine="text-davinci-002",
        prompt=prompt,
        max_tokens=1000,
        n=1,
        stop=None,
        temperature=0.5,
    )
    
    return eval(response.choices[0].text.strip())

def pdf_to_excel_with_openai(pdf_path, output_excel_path):
    """Convert PDF to Excel using OpenAI API for data extraction."""
    # Extract text from PDF
    text = extract_text_from_pdf(pdf_path)
    
    # Process text with OpenAI API
    structured_data = process_text_with_openai(text)
    
    # Convert to DataFrame and export to Excel
    df = pd.DataFrame(structured_data)
    df.to_excel(output_excel_path, index=False)
    print(f"Successfully extracted data to {output_excel_path}")

if __name__ == "__main__":
    pdf_path = r"C:\Users\ZANG\OneDrive\Desktop\DataExtraction\extracted_pages.pdf"
    output_excel_path = "data_extracted_openai.xlsx"
    pdf_to_excel_with_openai(pdf_path, output_excel_path)
