import pdfplumber
import re
import os
import json
import shutil
from pathlib import Path
import pandas as pd

def sanitize_filename_for_windows(filename):
    sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
    return sanitized

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
                    text += page_text + "\n"  # Ensure we separate pages with a newline
        print("Full extracted text from PDF:")
        print(text)
    except Exception as e:
        print(f"Error reading {pdf_file_path}: {e}")
    
    return text

def extract_invoice_number(text):
    patterns = [
        r'Rechnungsnr\.?\s*:\s*([\w\d-]+)', 
        r'Rechnung\s+(\d{4}/\d{4})', 
        r'(?:Rechnung\s*Nr\.?|Rechnungs-Nr\.?|Rechnungsnummer)[\s:]*[-\s]*([\w\d-]+)',  
        r'[\D]*(\d{8,})[\D]*Rechnungsnummer',  
        r'(\d{8,})\s*[\s\S]*?Rechnungsnummer\s*[:\s]*',  
        r'(\d{8,})\s*Rechnungsnummer\s*[:\s]*'
    ]
    
    for pattern in patterns:
        match = re.search(pattern, text, re.MULTILINE | re.DOTALL)
        if match:
            print(f"Found invoice number: {match.group(1).strip()}")  # Debugging output for found invoice number
            return match.group(1).strip()  # Return the first matched invoice number
    
    print("No invoice number found.")  # Debugging output when no invoice number is found
    return None

def get_files_in_folder(folder_path):
    files = []
    for root, dirs, filenames in os.walk(folder_path):
        for filename in filenames:
            if filename.lower().endswith('.pdf'):
                files.append(os.path.join(root, filename))
    return files

def extract_invoices_from_folder(folder_path):
    invoices = []
    files = get_files_in_folder(folder_path)

    for file in files:
        text = extract_text_from_pdf(file)
        invoice_number = extract_invoice_number(text)
        if invoice_number:
            invoices.append((file, invoice_number))
        else:
            invoices.append((file, 'No invoice number found'))

    return invoices

def display_invoices(invoices, filename_width=20):
    print(f"{'Filename':<{filename_width}} {'Rechnungsnummer':<25}")
    print("=" * (filename_width + 25))
    for filename, invoice_number in invoices:
        filename_base = os.path.basename(filename)
        if len(filename_base) > filename_width:
            filename_base = filename_base[:filename_width - 3] + "..."
        print(f"{filename_base:<{filename_width}} {invoice_number:<25}")

def rename_and_move_files(invoices, folder_selected):
    re_erledigt_path = folder_selected.replace('re_', 'Re_Erledigt')

    if not os.path.exists(re_erledigt_path):
        os.makedirs(re_erledigt_path)
        print(f"Created folder: {re_erledigt_path}")

    moved_files_info = []  # Store information about moved files

    for file_path, invoice_number in invoices:
        if invoice_number and invoice_number != 'No invoice number found':
            original_name = os.path.basename(file_path)
            new_invoice_number = invoice_number.replace('/', '-')  # Replace '/' with '-'
            new_name = f"{new_invoice_number}_{sanitize_filename_for_windows(original_name)}"
            new_path = os.path.join(os.path.dirname(file_path), new_name)

            try:
                os.rename(file_path, new_path)  # Rename the file
                destination_path = os.path.join(re_erledigt_path, new_name)

                print(f"Moving {new_name} to {destination_path}")  # Print the destination path
                shutil.move(new_path, destination_path)  # Move the renamed file to the new folder
                
                # Append information about the moved file
                moved_files_info.append({'filename': new_name, 'location': re_erledigt_path, 'status': 'moved'})
            except Exception as e:
                print(f"Error processing {file_path}: {e}")  # Handle any errors

    return moved_files_info

def update_excel_file(folder_selected, moved_files_info):
    excel_path = os.path.join(folder_selected, 'email_info.xlsx')

    if os.path.exists(excel_path):
        df = pd.read_excel(excel_path)
    else:
        df = pd.DataFrame(columns=['filename', 'location', 'status'])

    # Create a new DataFrame from the moved files info
    moved_df = pd.DataFrame(moved_files_info)

    # Concatenate the old and new DataFrames
    df = pd.concat([df, moved_df], ignore_index=True)

    # Write the updated DataFrame back to Excel
    df.to_excel(excel_path, index=False)

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
    folder_path = load_folder_info(json_path)

    if folder_path:
        folder_path = Path(folder_path)

        if folder_path.exists() and folder_path.is_dir():
            print(f"Processing files in the folder: {folder_path}")
            invoices = extract_invoices_from_folder(folder_path)
            display_invoices(invoices)
            moved_files_info = rename_and_move_files(invoices, str(folder_path))  # Get info of moved files
            update_excel_file(str(folder_path), moved_files_info)  # Update Excel file with moved files info
        else:
            print(f"The folder does not exist or is not a directory: {folder_path}")
    else:
        print("Could not retrieve the selected folder from the JSON file.")
