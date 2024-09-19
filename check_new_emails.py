import win32com.client
import pythoncom
from pathlib import Path
import re

print(pythoncom.__file__)


# This is working but only for the default email account outlook

class NewMailHandler:
    def OnItemAdd(self, item):
        try:
            # Check if the item is a mail item
            if item.Class == 43:  # olMailItem
                attachments = item.Attachments

                # Iterate through attachments
                for attachment in attachments:
                    # Check if the attachment is a PDF
                    if attachment.FileName.lower().endswith('.pdf'):
                        # Create a safe filename
                        filename = re.sub(r'[^0-9a-zA-Z\.]+', '', attachment.FileName)

                        # Save the PDF to the 're_' folder
                        attachment.SaveAsFile(re_dir / filename)
                        print(f"PDF saved: {filename}")

        except Exception as e:
            print(f"Error processing new email: {e}")

# Set up the output folder for PDFs
re_dir = Path(r"C:\Users\MaxEDV\Desktop\re_")
re_dir.mkdir(parents=True, exist_ok=True)

# Connect to Outlook and get the Inbox folder
try:
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # olFolderInbox
    print("Connected to Outlook Inbox.")
except Exception as e:
    print(f"Error connecting to Outlook: {e}")
    exit()

# Set up the event handler for the Inbox
try:
    items = inbox.Items
    print("Setting up event handler...")
    event_handler = win32com.client.WithEvents(items, NewMailHandler)
    print("Monitoring for new emails...")
except Exception as e:
    print(f"Error setting up event handler: {e}")
    exit()

# Keep the script running to listen for events
while True:
    try:
        # Process waiting COM messages
        pythoncom.PumpWaitingMessages()
    except KeyboardInterrupt:
        print("Script interrupted by user.")
        break
    except Exception as e:
        print(f"Error in message processing loop: {e}")
