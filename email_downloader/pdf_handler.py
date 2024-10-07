import PyPDF2
import logging
import os
import time
from pathlib import Path

# Function to merge PDF attachments and delete original PDFs
def merge_email_attachments(pdf_files, output_filename):
    merger = PyPDF2.PdfMerger()
    merged_file_path = None
    
    try:
        for pdf_file in pdf_files:
            # Convert Path objects to strings if necessary
            pdf_file_str = str(pdf_file) if isinstance(pdf_file, Path) else pdf_file
            
            if pdf_file_str.endswith('.pdf'):
                merger.append(pdf_file_str)
                logging.info(f"Added {pdf_file_str} to the merger.")
        
        if pdf_files:
            merged_file_path = os.path.join(os.path.dirname(str(pdf_files[0])), output_filename)
            with open(merged_file_path, 'wb') as f_out:
                merger.write(f_out)
            logging.info(f"Merged PDF saved as: {merged_file_path}")
            
            # Close the merger to release file handles
            merger.close()
            
            # Give the system a moment to release the file handles
            time.sleep(1)
            
            # Delete the original PDF files
            for pdf_file in pdf_files:
                pdf_file_path = Path(pdf_file) if not isinstance(pdf_file, Path) else pdf_file
                if pdf_file_path.exists() and pdf_file_path.suffix == '.pdf':
                    try:
                        pdf_file_path.unlink()  # This will delete the file
                        logging.info(f"Deleted original PDF: {pdf_file_path}")
                    except Exception as e:
                        logging.error(f"Error deleting PDF {pdf_file_path}: {e}")
        else:
            logging.warning("No PDF files to merge.")
    
    except Exception as e:
        logging.error(f"Error merging PDFs: {e}")
    
    finally:
        # Ensure that the merger is closed in case of an error
        merger.close()

    return merged_file_path
