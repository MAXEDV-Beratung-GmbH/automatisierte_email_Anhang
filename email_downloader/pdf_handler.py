import PyPDF2
import logging
import os
import time
from pathlib import Path

# Set to store processed files and avoid reprocessing
processed_files = set()

# Function to merge PDF attachments and delete original PDFs
def merge_email_attachments(pdf_files, output_filename):
    """Merges multiple PDF files into a single PDF and deletes the original files."""
    merger = PyPDF2.PdfMerger()
    merged_file_path = None
    
    try:
        for pdf_file in pdf_files:
            pdf_file_path = Path(pdf_file)  # Ensure we're using Path
            
            if pdf_file_path.suffix == '.pdf':
                # Check if the file has already been processed
                if str(pdf_file_path) not in processed_files:
                    merger.append(str(pdf_file_path))
                    logging.info(f"Added {pdf_file_path} to the merger.")
                    processed_files.add(str(pdf_file_path))  # Add to processed files
                else:
                    logging.info(f"File {pdf_file_path} already processed, skipping.")
        
        if pdf_files:
            merged_file_path = pdf_file_path.parent / output_filename
            with open(merged_file_path, 'wb') as f_out:
                merger.write(f_out)
            logging.info(f"Merged PDF saved as: {merged_file_path}")
        else:
            logging.warning("No PDF files to merge.")
    
    except Exception as e:
        logging.error(f"Error merging PDFs: {e}")
    
    finally:
        # Close the merger to release file handles
        merger.close()
        
        # Delete the original PDF files
        for pdf_file in pdf_files:
            pdf_file_path = Path(pdf_file)
            if pdf_file_path.exists() and pdf_file_path.suffix == '.pdf':
                try:
                    pdf_file_path.unlink()  # This will delete the file
                    logging.info(f"Deleted original PDF: {pdf_file_path}")
                except Exception as e:
                    logging.error(f"Error deleting PDF {pdf_file_path}: {e}")

    return merged_file_path
