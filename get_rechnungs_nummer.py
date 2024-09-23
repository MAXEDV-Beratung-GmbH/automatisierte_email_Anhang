import pdfplumber
import re
import os
import json
from pathlib import Path

# Extract text from a given PDF file
def extract_text_from_pdf(pdf_file_path):
    text = ''
    if not os.path.isfile(pdf_file_path):
        print(f"File does not exist: {pdf_file_path}")
        return text
    
    try:
        with pdfplumber.open(pdf_file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text
    except Exception as e:
        print(f"Error reading {pdf_file_path}: {e}")
    
    return text

# Extract invoice number from the extracted text
def extract_invoice_number(text):
    patterns = [
        r'Rechnungsnr\.\s*:\s*(RE\d+)', 
        r'Rechnung\s+(\d{4}/\d{4})', 
        r'(?:Rechnung\s*Nr\.?|Rechnungs-Nr\.?|Rechnungsnummer)[\s:]*[-\s]*([\w\d-]+)',  
        r'(\d{8,})\s*[\s\S]*?Rechnungsnummer\s*[:\s]*',  
        r'(\d{8,})\s*Rechnungsnummer\s*[:\s]*',  
        r'(\d{8,})\s*[\s\S]*?Rechnungsnummer'
    ]
    
    lines = text.split('\n')
    
    for i, line in enumerate(lines):
        if 'Rechnungs-Nr.' in line:
            next_line = lines[i + 1] if i + 1 < len(lines) else ''
            pattern = r'\b(\d{6,})\b'  
            match = re.search(pattern, next_line)
            if match:
                return match.group(1).strip()
    
    for pattern in patterns:
        match = re.search(pattern, text, re.MULTILINE | re.DOTALL)
        if match:
            return match.group(1).strip()
    
    return None

# Get all PDF files from the folder
def get_files_in_folder(folder_path):
    files = []
    for root, dirs, filenames in os.walk(folder_path):
        for filename in filenames:
            if filename.lower().endswith('.pdf'):
                files.append(os.path.join(root, filename))
    return files

# Extract invoice numbers from all PDFs in the folder
def extract_invoices_from_folder(folder_path):
    invoices = []
    files = get_files_in_folder(folder_path)

    for file in files:
        text = extract_text_from_pdf(file)
        invoice_number = extract_invoice_number(text)
        if invoice_number:
            invoices.append((os.path.basename(file), invoice_number))
        else:
            invoices.append((os.path.basename(file), 'No invoice number found'))

    return invoices

# Display the extracted invoice data
def display_invoices(invoices, filename_width=20):
    print(f"{'Filename':<{filename_width}} {'Rechnungsnummer':<25}")
    print("=" * (filename_width + 25))
    for filename, invoice_number in invoices:
        if len(filename) > filename_width:
            filename = filename[:filename_width - 3] + "..."
        print(f"{filename:<{filename_width}} {invoice_number:<25}")

# Load folder path from JSON
def load_folder_info(json_path):
    try:
        with open(json_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            return data.get('folder_selected')
    except FileNotFoundError:
        print(f"Could not find the file: {json_path}")
    except json.JSONDecodeError:
        print("Error decoding the JSON file.")
    return None

if __name__ == "__main__":
    json_path = 'user_info.json'
    
    # Load folder path from JSON
    folder_path = load_folder_info(json_path)

    if folder_path:
        folder_path = Path(folder_path)
        
        # Check if the folder exists and is a directory
        if folder_path.exists() and folder_path.is_dir():
            print(f"Processing files in the folder: {folder_path}")
            
            # Extract invoices from the folder
            invoices = extract_invoices_from_folder(folder_path)
            
            # Display the results
            display_invoices(invoices)
        else:
            print(f"The folder does not exist or is not a directory: {folder_path}")
    else:
        print("Could not retrieve the selected folder from the JSON file.")  
