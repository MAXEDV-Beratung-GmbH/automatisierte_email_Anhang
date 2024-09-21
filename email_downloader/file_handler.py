import os
import re
import json
import pandas as pd
from pathlib import Path

# Function to clean filenames by removing unsafe characters
def sanitize_filename(filename):
    return re.sub(r'[^0-9a-zA-Z\.]+', '', filename)

# Function to save email information to Excel without duplicating columns
def save_email_info_to_excel(email_data, excel_file):
    # Define the fixed columns
    columns = ["Date", "Email", "Subject"]

    # Check if the Excel file already exists
    if os.path.exists(excel_file):
        # Load the existing data
        df = pd.read_excel(excel_file)

        # Check if all required columns are present, if not, add missing columns
        for col in columns:
            if col not in df.columns:
                df[col] = None  # Add missing column with None values
    else:
        # Create a new DataFrame with the predefined columns
        df = pd.DataFrame(columns=columns)

    # Convert the new email_data into a DataFrame
    new_data = pd.DataFrame([email_data])

    # Concatenate the old and new data
    df = pd.concat([df, new_data], ignore_index=True)

    # Ensure only the fixed columns are saved
    df = df[columns]

    # Save the updated DataFrame back to Excel
    df.to_excel(excel_file, index=False)

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

        # Append the new data
        data.append(email_data)

        # Save to the JSON file
        with open(json_file, "w") as f:
            json.dump(data, f, indent=4)
    except Exception as e:
        print(f"Error saving email info: {e}")

# Function to check for new PDF files in the selected directory
def check_new_files(directory):
    try:
        existing_files = set(os.listdir(directory))
        return existing_files
    except Exception as e:
        print(f"Error checking files: {e}")
        return set()

# Create data directory if it doesn't exist
data_dir = Path("data")
data_dir.mkdir(exist_ok=True)

# Example usage
email_data = {
    "Date": "2024-09-19",
    "Email": "paniagua.ian.de@gmail.com", 
    "Subject": "Invoice Details"
}

# Set the path for the Excel file in the data directory
excel_file = data_dir / "email_info.xlsx"
json_file = data_dir / "email_info.json"

save_email_info_to_excel(email_data, excel_file)
save_email_info(email_data, json_file)
