import win32com.client
from pathlib import Path
import re
import pythoncom  # Needed for COM event handling
import json
from datetime import datetime

class NewMailHandler:
    def OnItemAdd(self, item):
        try:
            # Check if the item is a mail item
            if item.Class == 43:  # olMailItem
                attachments = item.Attachments

                # Initialize list to store JSON data
                json_data = []

                # Iterate through attachments
                for attachment in attachments:
                    # Check if the attachment is a PDF
                    if attachment.FileName.lower().endswith('.pdf'):
                        # Create a safe filename
                        filename = re.sub(r'[^0-9a-zA-Z\.]+', '', attachment.FileName)

                        # Save the PDF to the 're_' folder
                        attachment.SaveAsFile(pdf_dir / filename)
                        print(f"PDF saved: {filename}")

                        # Collect email information
                        sender = item.SenderEmailAddress
                        received_datetime = item.ReceivedTime

                        # Extract date and time separately
                        received_date = received_datetime.strftime('%Y-%m-%d')
                        received_time = received_datetime.strftime('%H:%M:%S')

                        # Add information to JSON data
                        json_data.append({
                            "sender": sender,
                            "received_date": received_date,
                            "received_time": received_time,
                            "pdf_name": filename
                        })

                # Save the data to a JSON file
                if json_data:  # Check if there is data to save
                    self.save_to_json(json_data)

        except Exception as e:
            print(f"Error processing new email: {e}")

    def save_to_json(self, data):
        # Path to the JSON file in the project directory
        json_file_path = project_dir / 'attachment_data.json'

        # Read existing data from the JSON file
        if json_file_path.exists():
            with open(json_file_path, 'r') as json_file:
                existing_data = json.load(json_file)
        else:
            existing_data = []

        # Append new data
        existing_data.extend(data)

        # Write updated data to the JSON file
        with open(json_file_path, 'w') as json_file:
            json.dump(existing_data, json_file, indent=4)

# Set up the output folder for PDFs
pdf_dir = Path(r"C:\Users\Teilnehmer\OneDrive - BBQ - Baumann Bildung und Qualifizierung GmbH\Desktop\re_")
pdf_dir.mkdir(parents=True, exist_ok=True)

# Set up the project directory for JSON file
project_dir = Path(r"C:\Users\Teilnehmer\OneDrive - BBQ - Baumann Bildung und Qualifizierung GmbH\Desktop\IMAP-OUTLOOK")
project_dir.mkdir(parents=True, exist_ok=True)

# Connect to Outlook and get the Inbox folder
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # olFolderInbox

# Debugging: Check connection to Outlook
print("Connected to Outlook Inbox.")

# Set up the event handler for the Inbox
items = inbox.Items

# Debugging: Ensure event handler is set
print("Setting up event handler...")

event_handler = win32com.client.WithEvents(items, NewMailHandler)

print("Monitoring for new emails...")

# Keep the script running to listen for events
while True:
    # Process waiting COM messages
    pythoncom.PumpWaitingMessages()
