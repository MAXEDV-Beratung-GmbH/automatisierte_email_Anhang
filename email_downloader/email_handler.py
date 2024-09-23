import imaplib
import email
from email.header import decode_header
from file_handler import save_email_info, save_email_info_to_excel, sanitize_filename, check_new_files
import os
import datetime


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

        new_files = check_new_files(re_dir)
        if new_files:
            print("New files detected: ", ', '.join(new_files))
        else:
            print("No new files found.")

        mail.select("inbox")
        status, messages = mail.search(None, '(UNSEEN)')
        mail_ids = messages[0].split()

        if mail_ids:
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
                                    email_data["Attachments"].append(filename)

                        # Save email information to JSON and Excel after handling attachments
                        save_email_info(email_data, json_file)
                        save_email_info_to_excel(email_data, excel_file)
        else:
            print("No new emails.")

    except Exception as e:
        print(f"Error checking inbox: {e}")