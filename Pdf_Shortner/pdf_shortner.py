from PyPDF2 import PdfReader, PdfWriter

def extract_pages(input_pdf_path, output_pdf_path, pages_to_extract):
    """
    Extracts specific pages from a PDF and saves them to a new PDF.

    Args:
    input_pdf_path (str): Path to the input PDF file.
    output_pdf_path (str): Path where the output PDF will be saved.
    pages_to_extract (list): List of page numbers to extract (0-indexed).
    """
    # Open the input PDF file
    with open(input_pdf_path, 'rb') as input_file:
        reader = PdfReader(input_file)
        writer = PdfWriter()

        # Extract the specified pages
        for page_num in pages_to_extract:
            if page_num < len(reader.pages):
                writer.add_page(reader.pages[page_num])
            else:
                print(f"Page number {page_num} is out of range.")

        # Write the extracted pages to a new PDF file
        with open(output_pdf_path, 'wb') as output_file:
            writer.write(output_file)

    print(f"Extracted pages saved to {output_pdf_path}.")

# Example usage
input_pdf_path = r'C:\Users\ZANG\OneDrive\Desktop\DataExtraction\SolarData.pdf' # Change directory accordingly
output_pdf_path = 'extracted_page291.pdf'
pages_to_extract = [291]  # Pages to extract (0-indexed)
extract_pages(input_pdf_path, output_pdf_path, pages_to_extract)
