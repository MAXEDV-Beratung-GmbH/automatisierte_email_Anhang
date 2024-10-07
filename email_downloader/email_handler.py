import imaplib
import email
from email.header import decode_header
from file_handler import save_email_info, save_email_info_to_excel, sanitize_filename, check_new_files
from pdf_handler import merge_email_attachments
import os
from datetime import datetime
import pdfplumber
import re
from pathlib import Path
import pandas as pd
import shutil
import logging
import json


# Basic logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# Function to connect to the IMAP server
def connect_imap(server, email_user, email_pass):
    try:
        mail = imaplib.IMAP4_SSL(server, port=993)
        mail.login(email_user, email_pass)
        logging.info("Login successful!")
        return mail
    except imaplib.IMAP4.error as e:
        logging.error(f"Error connecting: {e}")
        return None


# Function to extract the email address from the "From" field
def extract_email(from_field):
    match = re.search(r'<(.+?)>', from_field)
    return match.group(1) if match else from_field


# Function to sanitize file names on Windows
def sanitize_filename_for_windows(filename):
    return re.sub(r'[<>:"/\\|?*]', '_', filename)


# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_file_path):
    text = ''
    if not pdf_file_path.endswith('.pdf'):
        logging.info(f"Skipping non-PDF file: {pdf_file_path}")
        return text

    if not os.path.isfile(pdf_file_path):
        logging.error(f"File does not exist: {pdf_file_path}")
        return text

    try:
        with pdfplumber.open(pdf_file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        logging.info(f"Extracted text from PDF: {pdf_file_path}")
    except Exception as e:
        logging.error(f"Error reading {pdf_file_path}: {e}")

    return text


# Function to extract the invoice number from the text
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
            logging.info(f"Found invoice number: {match.group(1).strip()}")
            return match.group(1).strip()

    logging.info("No invoice number found.")
    return None

def get_files_in_folder(re_dir):
    files = []
    for root, dirs, filenames in os.walk(re_dir):
        for filename in filenames:
            if filename.lower().endswith('.pdf'):
                files.append(os.path.join(root, filename))
    return files

def extract_invoices_from_folder(re_dir):
    invoices = []
    files = get_files_in_folder(re_dir)

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

# Function to rename and move files
def rename_and_move_files(invoices, re_dir):
    re_erledigt_path = re_dir.replace('re_', 'Re_Erledigt')
    print('re_erledigt_path:' + re_erledigt_path)

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

# Function to check the inbox and download attachments
def check_inbox(mail, re_dir, json_file):
    try:
        excel_file = re_dir / "email_info.xlsx"

        # Check if there are new files in the directory
        new_files = check_new_files(re_dir)
        print(re_dir)
        invoices = []  # List to store processed invoices

        if new_files:
            logging.info(f"New files detected: {', '.join(map(str, new_files))}")
            for file in new_files:
                sanitized_filename = sanitize_filename(file)
                pdf_file_path = re_dir / sanitized_filename

                # Extract text from the PDF
                extracted_text = extract_text_from_pdf(str(pdf_file_path))
                print(extracted_text)
                invoice_number = extract_invoice_number(extracted_text)
                print(invoice_number)

                if invoice_number:
                    logging.info(f"Renaming and moving the file for invoice number: {invoice_number}")
                    invoices.append((pdf_file_path, invoice_number))
                    moved_files_info = rename_and_move_files(invoices, str(re_dir))
                    print('moved_files_info: ' + str(moved_files_info))
        else:
            logging.info("No new files found.")

        # Read unread emails
        mail.select("inbox")
        status, messages = mail.search(None, '(UNSEEN)')
        mail_ids = messages[0].split()

        if mail_ids:
            email_dates = []
            downloaded_files = []

            for mail_id in mail_ids:
                status, msg_data = mail.fetch(mail_id, "(RFC822)")
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_bytes(response_part[1])
                        subject, encoding = decode_header(msg["Subject"])[0]
                        if isinstance(subject, bytes):
                            subject = subject.decode(encoding if encoding else "utf-8")

                        date = msg.get("Date")
                        sender = extract_email(msg.get("From"))

                        email_date = datetime.strptime(date, '%a, %d %b %Y %H:%M:%S %z')
                        formatted_date = email_date.strftime('%Y-%m-%d')

                        logging.info(f"Processing email: {subject}")

                        email_data = {
                            "Date": date,
                            "Email": sender,
                            "Subject": subject,
                            "Attachments": []
                        }
                        pdf_attachments = []
                        if msg.is_multipart():
                            for part in msg.walk():
                                filename = part.get_filename()
                                if filename:
                                    filename = sanitize_filename(filename)
                                    filepath = re_dir / filename
                                    with open(filepath, "wb") as f:
                                        f.write(part.get_payload(decode=True))
                                    logging.info(f"Downloaded attachment: {filename}")
                                    email_data["Attachments"].append(filename)
                                    downloaded_files.append(filepath)
                                    pdf_attachments.append(filepath)
                        # Merge the PDFs if there are more than one
                        if len(pdf_attachments) > 1:
                            output_filename = f"merged_{formatted_date}_{sanitize_filename_for_windows(subject)}.pdf"
                            merged_file_path = merge_email_attachments(pdf_attachments, output_filename)
                            if merged_file_path:
                                downloaded_files.append(merged_file_path)
                        else:
                            downloaded_files.extend(pdf_attachments)
                        save_email_info(email_data, json_file)
                        save_email_info_to_excel(email_data, excel_file)

                        email_dates.append(formatted_date)

            return invoices, email_dates, downloaded_files
        
        if invoices:
            moved_files_info = rename_and_move_files(invoices, str(re_dir))
            if moved_files_info:
                logging.info("Files moved successfully.")
            else:
                logging.info("No files were moved.")
        else:
            logging.info("No invoices to move.")

    except Exception as e:
        logging.error(f"Error checking inbox: {e}")

    except Exception as e:
        logging.error(f"Error checking inbox: {e}")



# Call the necessary functions to process the inbox and move files
def main(server, email_user, email_pass, re_dir, json_file):
    mail = connect_imap(server, email_user, email_pass)

    if mail:
        invoices, email_dates, downloaded_files = check_inbox(mail, re_dir, json_file)
    else:
        logging.error("No connection to the IMAP server.")
        return

    invoices = extract_invoices_from_folder(re_dir)
    display_invoices(invoices)
    if invoices:
        rename_and_move_files(invoices, re_dir)  # Use folder_selected directly
    else:
        logging.info("No invoices to process.")

