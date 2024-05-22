import fitz  # PyMuPDF
import pandas as pd

def extract_text_from_pdf(pdf_path):
    # Open the PDF file
    document = fitz.open(pdf_path)
    all_text = []

    # Iterate through all the pages
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text = page.get_text("text")
        all_text.append(text)

    return all_text

def parse_tables(text_list):
    tables = []
    for text in text_list:
        lines = text.split('\n')
        data = [line.split() for line in lines if line.strip() != '']
        tables.append(pd.DataFrame(data))
    return tables

def pdf_to_excel(pdf_path, excel_path):
    # Extract text from the PDF
    text_list = extract_text_from_pdf(pdf_path)
    
    # Parse the extracted text into tables
    tables = parse_tables(text_list)
    
    # Create an Excel writer object
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        # Write each table to a separate sheet in the Excel file
        for i, table in enumerate(tables):
            table.to_excel(writer, sheet_name=f'Table_{i+1}', index=False)
            print(f"Table {i+1} written to sheet Table_{i+1}")
    
    print(f"Data has been written to {excel_path}")

# Example usage
pdf_path = 'path/to/your/file.pdf'
excel_path = 'path/to/save/your/file.xlsx'
pdf_to_excel(pdf_path, excel_path)
