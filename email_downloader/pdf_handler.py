import PyPDF2
from pathlib import Path

def merge_pdfs(folder_path, output_filename):
    # Get all PDF files in the specified folder
    pdf_files = list(Path(folder_path).glob("*.pdf"))
    pdf_files.sort()  # Sort files if needed

    # Check if there are any PDF files to merge
    if not pdf_files:
        print("No PDF files found to merge.")
        return

    # Define the output path for the merged PDF
    output_path = Path(folder_path) / output_filename

    # Create a PDF writer object
    with open(output_path, "wb") as output_file:
        pdf_writer = PyPDF2.PdfWriter()
        for pdf_file in pdf_files:
            with open(pdf_file, "rb") as f:
                pdf_reader = PyPDF2.PdfReader(f)
                # Add each page of the PDF file to the writer
                for page in range(len(pdf_reader.pages)):
                    pdf_writer.add_page(pdf_reader.pages[page])
            print(f"Merged: {pdf_file.name}")

        # Write the merged PDF to the output file
        pdf_writer.write(output_file)

    print(f"Merged PDF saved as: {output_path}")
