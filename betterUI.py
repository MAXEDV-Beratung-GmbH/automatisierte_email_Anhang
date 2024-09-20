import imaplib
import email
from email.header import decode_header
import time
import os
import re
import json
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, ttk
import threading

# Function to select the folder where PDFs will be saved
def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide root window
    folder_selected = filedialog.askdirectory(title="Select folder to save PDFs")
    if not folder_selected:
        print("No folder selected. Exiting.")
        exit()
    return Path(folder_selected)

# Function to clean filenames by removing unsafe characters
def sanitize_filename(filename):
    return re.sub(r'[^0-9a-zA-Z\.]+', '', filename)

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

# Function to save email information to JSON
def save_email_info(email_data, json_file):
    try:
        if os.path.exists(json_file):
            with open(json_file, "r") as f:
                try:
                    data = json.load(f)
                except json.JSONDecodeError:
                    print("Error decoding JSON, starting fresh.")
                    data = []
        else:
            data = []

        data.append(email_data)

        with open(json_file, "w") as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        print(f"Error saving email info: {e}")
        
#download all attachments
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
                                print(f"Content-Disposition: {content_disposition}")  # Debugging line
                                if "attachment" in content_disposition:
                                    filename = part.get_filename()
                                    if filename:
                                        filename = sanitize_filename(filename)
                                        filepath = re_dir / filename
                                        print(f"Saving to: {filepath}")  # Print the full file path
                                        try:
                                            with open(filepath, "wb") as f:
                                                f.write(part.get_payload(decode=True))
                                            print(f"Attachment saved: {filename}")
                                        except Exception as e:
                                            print(f"Error saving attachment {filename}: {e}")
                                    else:
                                        print("Attachment filename is empty.")
                        else:
                            print(f"No attachments in email: {subject}")
        else:
            print("No new emails.")
    
    except Exception as e:
        print(f"Error checking inbox: {e}")

# Function to start checking inbox in a separate thread
def start_checking_inbox(mail, re_dir, json_file):
    try:
        while True:
            check_inbox(mail, re_dir, json_file)
            print("Waiting 30 seconds for the next check...")
            time.sleep(30)
    except KeyboardInterrupt:
        print("Exiting script.")
    finally:
        mail.logout()

# Function to start the application
def start_app():
    root = tk.Tk()
    root.title("Email PDF Downloader")
    
    # Get screen width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    # Set window size
    window_width = 300
    window_height = 150
    
    # Calculate position for centering
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)

    # Set geometry
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    tk.Label(root, text="Email:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
    email_entry = tk.Entry(root, width=30)
    email_entry.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(root, text="Password:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    password_entry = tk.Entry(root, show='*', width=30)
    password_entry.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(root, text="Email Provider:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
    provider_combobox = ttk.Combobox(root, values=["IONOS", "Outlook", "Gmail", "Yahoo"])
    provider_combobox.grid(row=2, column=1, padx=5, pady=5)

    def submit():
        email_user = email_entry.get()
        email_pass = password_entry.get()
        provider = provider_combobox.get()

        provider_map = {
            "IONOS": "imap.ionos.es",
            "Outlook": "outlook.office365.com",
            "Gmail": "imap.gmail.com",
            "Yahoo": "imap.mail.yahoo.com"
        }
        server = provider_map.get(provider)

        if not server:
            print("Invalid provider selected.")
            return

        mail = connect_imap(server, email_user, email_pass)

        if mail:
            re_dir = select_folder()

            # Open a new window to display success message
            success_window = tk.Toplevel(root)
            success_window.title("Connection Successful")
            success_window.geometry("300x100")
            success_label = tk.Label(success_window, text=f"Connection established successfully.\nEmail: {email_user}")
            success_label.pack(pady=20)

            # Start checking inbox in a new thread
            threading.Thread(target=start_checking_inbox, args=(mail, re_dir, Path("data/email_info.json")), daemon=True).start()

            # Hide the main window
            root.withdraw()

    # Submit button
    submit_button = tk.Button(root, text="Start", command=submit)
    submit_button.grid(row=3, columnspan=2, pady=10)

    root.mainloop()

if __name__ == "__main__":
    start_app()