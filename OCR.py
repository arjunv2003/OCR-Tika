import os
from pdf2image import convert_from_path
import pytesseract
from tika import parser
from docx import Document

def extract_text_from_pdf(file_path):
    """Extract text from a text-based PDF using Tika."""
    parsed = parser.from_file(file_path)
    return parsed["content"]

def extract_text_from_scanned_pdf(file_path):
    """Extract text from a scanned PDF using Tesseract OCR."""
    images = convert_from_path(file_path, dpi=300)  # Higher DPI improves OCR accuracy
    extracted_text = ""
    for i, image in enumerate(images):
        print(f"Processing page {i + 1} of scanned PDF...")
        text = pytesseract.image_to_string(image, lang='eng')  # Specify language if needed
        extracted_text += text + "\n"
    return extracted_text

def extract_text_from_docx(file_path):
    """Extract text from a DOCX file using python-docx."""
    doc = Document(file_path)
    return "\n".join([paragraph.text for paragraph in doc.paragraphs])

def extract_text_from_doc(file_path):
    """Extract text from a DOC file using Tika."""
    parsed = parser.from_file(file_path)
    return parsed["content"]

def process_cv(file_path):
    """Process a single CV based on its file type."""
    if file_path.lower().endswith('.pdf'):
        # Check if the PDF is scanned or text-based
        parsed = parser.from_file(file_path)
        if parsed["content"] is None or parsed["content"].strip() == "":
            print(f"{file_path} is a scanned PDF. Using OCR...")
            return extract_text_from_scanned_pdf(file_path)
        else:
            print(f"{file_path} is a text-based PDF. Using Tika...")
            return extract_text_from_pdf(file_path)
    elif file_path.lower().endswith('.docx'):
        print(f"{file_path} is a DOCX file. Using python-docx...")
        return extract_text_from_docx(file_path)
    elif file_path.lower().endswith('.doc'):
        print(f"{file_path} is a DOC file. Using Tika...")
        return extract_text_from_doc(file_path)
    else:
        raise ValueError(f"Unsupported file type: {file_path}")

def save_text_to_file(text, output_folder, filename):
    """Save the extracted text to a .txt file."""
    output_file_path = os.path.join(output_folder, f"{os.path.splitext(filename)[0]}.txt")
    with open(output_file_path, "w", encoding="utf-8") as file:
        file.write(text)
    print(f"Saved extracted text to {output_file_path}")

def process_all_cvs(folder_path, output_folder):
    """Process all CVs in the given folder and save the extracted text to separate files."""
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)  # Create the output folder if it doesn't exist

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            print(f"\nProcessing file: {filename}")
            try:
                extracted_text = process_cv(file_path)
                save_text_to_file(extracted_text, output_folder, filename)
                print(f"Successfully extracted and saved text from {filename}")
            except Exception as e:
                print(f"Error processing {filename}: {e}")

# Folder containing the CVs
cv_folder = r"C:\Users\USER\Desktop\windows\pdf_image"  # Replace with the actual folder path
output_folder = r"C:\Users\USER\Desktop\windows\output"   # Replace with the desired output folder path

# Process all CVs in the folder and save the extracted text
process_all_cvs(cv_folder, output_folder)
