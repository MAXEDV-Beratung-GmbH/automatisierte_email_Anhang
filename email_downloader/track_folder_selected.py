import os
import time
import json

def load_user_info():
    # Load user information from the JSON file
    with open("user_info.json", "r") as f:
        return json.load(f)

def check_for_new_files(folder_path, processed_files):
    # Get the current files in the directory
    current_files = set(os.listdir(folder_path))
    # Find new files by subtracting processed files from current files
    new_files = current_files - processed_files

    return new_files

def main():
    # Load user info to get the folder path
    user_info = load_user_info()
    folder_path = user_info["folder_selected"]

    # Set to keep track of processed files
    processed_files = set(os.listdir(folder_path))

    print(f"Monitoring folder: {folder_path}")

    while True:
        # Check for new files
        new_files = check_for_new_files(folder_path, processed_files)

        if new_files:
            print(f"New files are: {', '.join(new_files)}")
            # Update the processed files set with new files
            processed_files.update(new_files)

        # Wait some time before checking again
        time.sleep(10)  # Change the time if necessary

if __name__ == "__main__":
    main()
