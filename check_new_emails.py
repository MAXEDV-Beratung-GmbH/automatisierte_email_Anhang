import win32com.client
from pathlib import Path
import re
import pythoncom  # Needed for COM event handling

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
# re_dir = Path.cwd() / "re_"  (project folder)

# save in the PC
# desktop = Path.home() / "Desktop"
# re_dir = desktop / "re_"
re_dir = Path(r"C:\Users\Teilnehmer\OneDrive - BBQ - Baumann Bildung und Qualifizierung GmbH\Desktop\re_")
re_dir.mkdir(parents=True, exist_ok=True)



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
