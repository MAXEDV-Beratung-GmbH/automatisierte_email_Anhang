import imaplib
import email
from email.header import decode_header
import re
from file_handler import save_email_info, sanitize_filename


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

# Function to download all attachments from unread emails
def check_inbox(mail, re_dir, json_file):
    try:
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
                        sender, encoding = decode_header(msg["From"])[0]
                        if isinstance(sender, bytes):
                            sender = sender.decode(encoding if encoding else "utf-8")

                        print(f"Processing email: {subject}")

                        email_data = {
                            "date": date,
                            "sender": sender,
                            "subject": subject
                        }
                        save_email_info(email_data, json_file)

                        if msg.is_multipart():
                            print("Email is multipart, checking attachments...")
                            for part in msg.walk():
                                content_disposition = str(part.get("Content-Disposition"))
                                if "attachment" in content_disposition:
                                    filename = part.get_filename()
                                    if filename:
                                        filename = sanitize_filename(filename)
                                        filepath = re_dir / filename
                                        with open(filepath, "wb") as f:
                                            f.write(part.get_payload(decode=True))
                                        print(f"Attachment saved: {filename}")
                                    else:
                                        print("Attachment filename is empty.")
        else:
            print("No new emails.")
    
    except Exception as e:
        print(f"Error checking inbox: {e}")
