import docx2txt
from docx import Document
from PyPDF2 import PdfReader
import re

# Function to process and clean the content of a DOCX file
def _clean_docx(docx_path, output_path="contract_to_txt_cleaned.txt"):
    """
    Clean and enhance text extracted from a DOCX file.
    """
    # Extract text from DOCX
    raw_text = docx2txt.process(docx_path)
    
    # Remove unwanted spaces and normalize
    cleaned_text = '\n'.join([line.strip() for line in raw_text.splitlines() if line.strip()])
    
    # Correct common OCR artifacts (if any)
    cleaned_text = re.sub(r'\bi\b', 'I', cleaned_text)
    cleaned_text = re.sub(r'\b1\b', 'I', cleaned_text)
    
    # Save to output
    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(cleaned_text)
    

# Function to process and clean the content of a PDF file
def _clean_pdf(pdf_path, output_path="contract_to_txt_cleaned.txt"):
    """
    Clean and enhance text extracted from a PDF file.
    """
    try:
        # Extract text using PyPDF2
        pdf_reader = PdfReader(pdf_path)
        extracted_text = ''
        
        for page in pdf_reader.pages:
            extracted_text += page.extract_text()
        
        # Fallback to OCR if PDF extraction fails
        if not extracted_text.strip():
            print("No text extracted; switching to OCR...")
            extracted_text = perform_ocr_on_pdf(pdf_path)
        
        # Clean the extracted text
        cleaned_text = '\n'.join([line.strip() for line in extracted_text.splitlines() if line.strip()])
        
        # Save to output
        with open(output_path, 'w', encoding='utf-8') as file:
            file.write(cleaned_text)
        return cleaned_text
        
    except Exception as e:
        print(f"Error in enhanced_clean_pdf: {e}")
        return None

# Function to convert a txt file to docx file
def txt_to_docx(txt_filepath, cleaned_docx_filepath):
    # Initialize a Document
    doc = Document()
    
    # Open and read the txt file
    with open(txt_filepath, 'r', encoding = "utf8") as txt_file:
        # Loop over each line in the txt file
        for line in txt_file:
            # Add a paragraph for each line in txt to docx
            doc.add_paragraph(line.strip())
            
    # Save the Document
    doc.save(cleaned_docx_filepath)

def convert_to_txt(contract_in):
    dot_pdf = '.pdf'
    dot_docx = '.docx'
    if contract_in.endswith(dot_pdf):
        _clean_pdf(contract_in)
    elif contract_in.endswith(dot_docx):
        _clean_docx(contract_in)
    else:
        return ('Error: unexpected file type')