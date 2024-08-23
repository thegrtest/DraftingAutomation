import os
from PIL import Image

def convert_tif_to_pdf(directory):
    # Iterate over all files in the directory
    for filename in os.listdir(directory):
        if filename.lower().endswith(".tif") or filename.lower().endswith(".tiff"):
            # Open the TIF file
            filepath = os.path.join(directory, filename)
            with Image.open(filepath) as img:
                # Convert the image to RGB mode (necessary for PDF conversion)
                img = img.convert("RGB")
                
                # Save as PDF with the same name as the TIF file
                pdf_filename = os.path.splitext(filename)[0] + ".pdf"
                pdf_filepath = os.path.join(directory, pdf_filename)
                img.save(pdf_filepath, "PDF")

            print(f"Converted {filename} to {pdf_filename}")

if __name__ == "__main__":
    # Get the current directory
    current_directory = os.path.dirname(os.path.abspath(__file__))
    
    # Convert TIF files in the current directory
    convert_tif_to_pdf(current_directory)
