import pdfplumber
import re
import os
import json
import shutil
from pathlib import Path

# Function to sanitize filenames to be compatible with Windows (removes invalid characters)
def sanitize_filename_for_windows(filename):
    sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
    return sanitized

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_file_path):
    text = ''
    if not os.path.isfile(pdf_file_path):  # Check if the file exists
        print(f"File does not exist: {pdf_file_path}")
        return text
    
    try:
        with pdfplumber.open(pdf_file_path) as pdf:
            for page in pdf.pages:  # Iterate through all pages of the PDF
                page_text = page.extract_text()  # Extract text from each page
                if page_text:
                    text += page_text  # Accumulate text from all pages
    except Exception as e:
        print(f"Error reading {pdf_file_path}: {e}")  # Handle any errors while reading the file
    
    return text

# Function to extract the invoice number from the extracted text
def extract_invoice_number(text):
    patterns = [
        r'Rechnungsnr\.\s*:\s*(RE\d+)',  # Matches specific invoice number formats (like RE followed by digits)
        r'Rechnung\s+(\d{4}/\d{4})',  # Matches invoice number patterns with year-based structure
        r'(?:Rechnung\s*Nr\.?|Rechnungs-Nr\.?|Rechnungsnummer)[\s:]*[-\s]*([\w\d-]+)',  # Matches invoice number in different notations
        r'(\d{8,})\s*[\s\S]*?Rechnungsnummer\s*[:\s]*',  # Matches any sequence of digits
        r'(\d{8,})\s*Rechnungsnummer\s*[:\s]*',  
        r'(\d{8,})\s*[\s\S]*?Rechnungsnummer'
    ]
    
    # Try to match the text with one of the patterns above
    for pattern in patterns:
        match = re.search(pattern, text, re.MULTILINE | re.DOTALL)
        if match:
            return match.group(1).strip()  # Return the first matching group (the invoice number)
    
    return None

# Function to get all files in a folder (recursively searches subfolders)
def get_files_in_folder(folder_path):
    files = []
    for root, dirs, filenames in os.walk(folder_path):  # Walk through all subfolders
        for filename in filenames:
            files.append(os.path.join(root, filename))  # Store the full path of each file
    return files

# Function to extract invoice information from all files in a folder and link them to the corresponding email attachments
def extract_invoices_from_folder(folder_path, email_attachments):
    invoices = []
    files = get_files_in_folder(folder_path)  # Get all files in the folder

    for file in files:
        text = ''
        if file.lower().endswith('.pdf'):  # Process only PDF files
            text = extract_text_from_pdf(file)  # Extract text from the PDF file
            invoice_number = extract_invoice_number(text)  # Extract the invoice number from the text
            if invoice_number:
                invoices.append((file, invoice_number, email_attachments))  # Store the file, invoice number, and related email attachments
            else:
                invoices.append((file, 'No invoice number found', email_attachments))  # No invoice number found
        else:
            invoices.append((file, 'No invoice number found for non-PDF file', email_attachments))  # Non-PDF file

    return invoices

# Function to display the invoice details in the console
def display_invoices(invoices, filename_width=20):
    print(f"{'Filename':<{filename_width}} {'Rechnungsnummer':<25}")  # Column headers
    print("=" * (filename_width + 25))
    for filename, invoice_number, _ in invoices:
        filename_base = os.path.basename(filename)  # Extract the base filename
        if len(filename_base) > filename_width:  # Shorten long filenames for display
            filename_base = filename_base[:filename_width - 3] + "..."
        print(f"{filename_base:<{filename_width}} {invoice_number:<25}")

# Function to rename and move files with invoice numbers to a new folder
def rename_and_move_files(invoices, folder_selected):
    re_erledigt_path = folder_selected.replace('re_', 'Re_Erledigt')  # Set the target folder name

    if not os.path.exists(re_erledigt_path):  # Create the folder if it doesn't exist
        os.makedirs(re_erledigt_path)
        print(f"Created folder: {re_erledigt_path}")

    for file_path, invoice_number, email_attachments in invoices:
        if invoice_number != 'No invoice number found':
            invoice_number = invoice_number.replace('/', '-')  # Replace '/' with '-' in invoice numbers

            # Rename and move all attachments from the same email
            for attachment in email_attachments:
                attachment_path = os.path.join(folder_selected, attachment)
                if os.path.exists(attachment_path):  # Check if the attachment file exists
                    original_name = os.path.basename(attachment_path)  # Get the attachment's original name
                    new_name = f"{invoice_number}_{sanitize_filename_for_windows(original_name)}"  # Create the new filename
                    destination_path = os.path.join(re_erledigt_path, new_name)  # Define the new path

                    try:
                        shutil.move(attachment_path, destination_path)  # Move the attachment file
                        print(f"Moving {original_name} to {destination_path}")
                    except Exception as e:
                        print(f"Error processing {attachment_path}: {e}")  # Handle any errors during file operations

# Function to load user information (e.g., folder selection) from a JSON file
def load_user_info(json_path):
    try:
        with open(json_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            return data.get('folder_selected')  # Return the folder path
    except FileNotFoundError:
        print(f"Could not find the file: {json_path}")
    except json.JSONDecodeError:
        print("Error decoding the JSON file.")
    return None

# Function to load email information (e.g., attachments) from a JSON file
def load_email_info(json_path):
    try:
        with open(json_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
            return data  # Return all email information
    except FileNotFoundError:
        print(f"Could not find the file: {json_path}")
    except json.JSONDecodeError:
        print("Error decoding the JSON file.")
    return None

# Main execution starts here
if __name__ == "__main__":
    user_info_path = 'user_info.json'
    email_info_path = 'data/email_info.json'

    folder_path = load_user_info(user_info_path)  # Load the folder path from the user's info
    email_info = load_email_info(email_info_path)  # Load email information (e.g., attachments)

    if email_info:
        for email in email_info:
            email_attachments = email['Attachments']  # Get the attachments for each email
            print(f"Processing email: {email['Subject']}")
            invoices = extract_invoices_from_folder(folder_path, email_attachments)  # Extract invoices linked to attachments
            display_invoices(invoices)  # Display the invoice information
            rename_and_move_files(invoices, folder_path)  # Rename and move files as needed
    else:
        print("Could not retrieve email information.")  # Error handling if email info is missing
