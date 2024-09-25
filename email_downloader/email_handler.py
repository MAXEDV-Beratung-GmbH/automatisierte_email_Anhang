import imaplib
import email
from email.header import decode_header
from file_handler import save_email_info, save_email_info_to_excel, sanitize_filename, check_new_files
from pdf_handler import merge_pdfs  # Ensure pdf_handler.py is in the same directory
import os
from datetime import datetime  # Import for handling dates

# Function to connect to the IMAP server
def connect_imap(server, email_user, email_pass):
    try:
        mail = imaplib.IMAP4_SSL(server, port=993)
        mail.login(email_user, email_pass)
        print("Login successful!")
        return mail
    except imaplib.IMAP4.error as e:
        print(f"Error connecting: {e}")
        return None

# Function to extract the email address from the "From" field
def extract_email(from_field):
    import re
    match = re.search(r'<(.+?)>', from_field)
    return match.group(1) if match else from_field

# Function to download all attachments from unread emails
def check_inbox(mail, re_dir, json_file):
    try:
        excel_file = re_dir / "email_info.xlsx"

        # Check for new files in the directory
        new_files = check_new_files(re_dir)
        if new_files:
            print("New files detected: ", ', '.join(new_files))
        else:
            print("No new files found.")

        mail.select("inbox")
        status, messages = mail.search(None, '(UNSEEN)')
        mail_ids = messages[0].split()

        if mail_ids:
            email_dates = []  # List to store the dates of processed emails
            downloaded_files = []  # List to keep track of downloaded attachments

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

                        # Parse the date and format it
                        email_date = datetime.strptime(date, '%a, %d %b %Y %H:%M:%S %z')
                        formatted_date = email_date.strftime('%Y-%m-%d')  # Format: YYYY-MM-DD

                        print(f"Processing email: {subject}")

                        email_data = {
                            "Date": date,
                            "Email": sender,
                            "Subject": subject,
                            "Attachments": []  # Initialize as a list
                        }

                        if msg.is_multipart():
                            for part in msg.walk():
                                filename = part.get_filename()
                                if filename:
                                    filename = sanitize_filename(filename)
                                    filepath = re_dir / filename
                                    with open(filepath, "wb") as f:
                                        f.write(part.get_payload(decode=True))
                                    print(f"Downloaded attachment: {filename}")
                                    email_data["Attachments"].append(filename)  # Add the filename to the list
                                    downloaded_files.append(filepath)  # Track the downloaded file

                        # Save email information to JSON and Excel after handling attachments
                        save_email_info(email_data, json_file)
                        save_email_info_to_excel(email_data, excel_file)

                        # Add the formatted date to the list
                        email_dates.append(formatted_date)

            # After processing all emails, merge the downloaded PDFs
            if email_dates:
                last_email_date = email_dates[-1]  # Get the date of the last processed email
                merged_filename = f"{last_email_date}_merged.pdf"  # Create the combined filename
                merge_pdfs(re_dir, merged_filename)  # Specify the output filename for the merged PDF

                # Delete the downloaded files after merging
                for file in downloaded_files:
                    try:
                        os.remove(file)
                        print(f"Deleted file: {file}")
                    except Exception as e:
                        print(f"Error deleting file {file}: {e}")

            else:
                print("No emails with attachments were processed.")
        else:
            print("No new emails.")

    except Exception as e:
        print(f"Error checking inbox: {e}")
