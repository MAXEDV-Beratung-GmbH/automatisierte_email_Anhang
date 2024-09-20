# file_handler.py
import os
import re
import json
from pathlib import Path

# Function to clean filenames by removing unsafe characters
def sanitize_filename(filename):
    return re.sub(r'[^0-9a-zA-Z\.]+', '', filename)

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
