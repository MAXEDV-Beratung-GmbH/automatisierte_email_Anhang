# gui.py

import tkinter as tk
from tkinter import filedialog, ttk
from pathlib import Path
from config import EMAIL_PROVIDERS
from email_handler import connect_imap, check_inbox
from file_handler import save_email_info
import threading
import time

# Function to select the folder where PDFs will be saved
def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide root window
    folder_selected = filedialog.askdirectory(title="Select folder to save PDFs")
    if not folder_selected:
        print("No folder selected. Exiting.")
        exit()
    return Path(folder_selected)

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
    
    tk.Label(root, text="Email:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
    email_entry = tk.Entry(root, width=30)
    email_entry.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(root, text="Password:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
    password_entry = tk.Entry(root, show='*', width=30)
    password_entry.grid(row=1, column=1, padx=5, pady=5)

    tk.Label(root, text="Email Provider:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
    provider_combobox = ttk.Combobox(root, values=list(EMAIL_PROVIDERS.keys()))
    provider_combobox.grid(row=2, column=1, padx=5, pady=5)

    def submit():
        email_user = email_entry.get()
        email_pass = password_entry.get()
        provider = provider_combobox.get()

        server = EMAIL_PROVIDERS.get(provider)

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

            # Center the success window on the main window
            x = root.winfo_x() + (root.winfo_width() // 2) - 150  # 150 is half of the success window width
            y = root.winfo_y() + (root.winfo_height() // 2) - 50   # 50 is half of the success window height
            success_window.geometry(f"+{x}+{y}")

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
