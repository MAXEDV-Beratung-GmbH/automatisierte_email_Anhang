import imaplib
import email
from email.header import decode_header
from file_handler import save_email_info, save_email_info_to_excel, sanitize_filename, check_new_files
from pdf_handler import merge_pdfs
import os
from datetime import datetime
import pdfplumber
import re
from pathlib import Path
import pandas as pd
import shutil
import logging

# Configuración básica del logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


# Función para conectarse al servidor IMAP
def connect_imap(server, email_user, email_pass):
    try:
        mail = imaplib.IMAP4_SSL(server, port=993)
        mail.login(email_user, email_pass)
        logging.info("Login successful!")
        return mail
    except imaplib.IMAP4.error as e:
        logging.error(f"Error connecting: {e}")
        return None


# Función para extraer la dirección de correo del campo "From"
def extract_email(from_field):
    match = re.search(r'<(.+?)>', from_field)
    return match.group(1) if match else from_field


# Función para sanitizar los nombres de archivos en Windows
def sanitize_filename_for_windows(filename):
    return re.sub(r'[<>:"/\\|?*]', '_', filename)


# Función para extraer texto de un archivo PDF
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


# Función para extraer el número de factura de un texto
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


# Función para revisar la bandeja de entrada y descargar archivos adjuntos
# Función para revisar la bandeja de entrada y descargar archivos adjuntos
def check_inbox(mail, re_dir, json_file):
    try:
        excel_file = re_dir / "email_info.xlsx"

        # Verificar si hay nuevos archivos en el directorio
        new_files = check_new_files(re_dir)
        invoices = []  # Lista para almacenar facturas procesadas

        if new_files:
            logging.info(f"New files detected: {', '.join(map(str, new_files))}")
            for file in new_files:
                sanitized_filename = sanitize_filename(file)
                pdf_file_path = re_dir / sanitized_filename

                # Extraer texto del PDF
                extracted_text = extract_text_from_pdf(str(pdf_file_path))
                invoice_number = extract_invoice_number(extracted_text)

                if invoice_number:
                    logging.info(f"Renaming and moving the file for invoice number: {invoice_number}")
                    
                    # Renombrar y mover el archivo inmediatamente después de encontrar la factura
                    moved_file = rename_and_move_single_file(sanitized_filename, invoice_number, re_dir, json_file)
                    
                    # Si el archivo fue renombrado y movido correctamente, lo agregamos a la lista de facturas
                    if moved_file:
                        invoices.append(moved_file)
        else:
            logging.info("No new files found.")

        # Leer correos electrónicos no leídos
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

                        save_email_info(email_data, json_file)
                        save_email_info_to_excel(email_data, excel_file)

                        email_dates.append(formatted_date)

            return invoices, email_dates, downloaded_files

    except Exception as e:
        logging.error(f"Error checking inbox: {e}")

# Función para renombrar y mover archivos
def rename_and_move_files(invoices, re_dir, config):
    # Extraer el folder_selected desde la configuración
    folder_selected = config.get('folder_selected', str(re_dir))  # Si no se proporciona, usa re_dir por defecto
    re_erledigt_path = Path(folder_selected) / "Re_Erledigt"

    # Crear el directorio si no existe
    if not re_erledigt_path.exists():
        re_erledigt_path.mkdir(parents=True)
        logging.info(f"Created folder: {re_erledigt_path}")

    moved_files_info = []

    for invoice in invoices:
        file_path = re_dir / invoice['filename']
        invoice_number = invoice['invoice_number']

        logging.info(f"Processing file {file_path} with invoice number {invoice_number}")

        if invoice_number:
            new_invoice_number = invoice_number.replace('/', '-')
            new_name = f"{new_invoice_number}_{sanitize_filename_for_windows(invoice['filename'])}"
            new_path = re_dir / new_name

            try:
                # Renombrar el archivo
                logging.info(f"Renaming file {file_path} to {new_path}")
                file_path.rename(new_path)

                # Mover el archivo renombrado a la carpeta Re_Erledigt en la ruta seleccionada
                destination_path = re_erledigt_path / new_name
                logging.info(f"Moving file {new_name} to {destination_path}")
                shutil.move(new_path, destination_path)

                moved_files_info.append({'filename': new_name, 'location': str(re_erledigt_path), 'status': 'moved'})
            except Exception as e:
                logging.error(f"Error processing {file_path}: {e}")

    return moved_files_info

# Llamar las funciones necesarias para procesar bandeja de entrada y mover archivos
def main(server, email_user, email_pass, re_dir, json_file):
    mail = connect_imap(server, email_user, email_pass)
    
    # Cargar la configuración desde el archivo JSON
    with open(json_file, 'r') as f:
        config = json.load(f)

    if mail:
        invoices, email_dates, downloaded_files = check_inbox(mail, re_dir, json_file)
        if invoices:
            rename_and_move_files(invoices, re_dir, config)  # Pasar la configuración
        else:
            logging.info("No invoices to process.")
