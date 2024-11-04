from PyPDF2 import PdfReader
import mmap

# Converts pdf to txt
def pdf_to_txt(pdf_path):
    no_ext = pdf_path[:-4]
    txt_fname = no_ext + ".txt"

    pdf_file_obj = open(pdf_path, 'rb')

    pdf_reader = PdfReader(pdf_file_obj)

    page_count = len(pdf_reader.pages)

    page_obj = pdf_reader.pages[page_count - 1]

    pdf_text = ''

    for i in range(page_count - 1):
        page_obj = pdf_reader.pages[i]
        pdf_text += (page_obj.extract_text())

    with open(txt_fname, 'w', encoding="utf-8") as file:
        file.write(pdf_text)
    
    return txt_fname


# Method that searches txt files for things in findings_list
def flag_findings(txt_file):
    findings_list = [b"Material weaknesses identified? Yes",
                     b"Noncompliance material to financial statements noted? Yes",
                     b"Significant deficiencies identified? Yes",
                    ]
    with open(txt_file) as f:
        for i in findings_list:
            s = mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ)
            if s.find(i) != -1:
                return True
    return False


# Main function that takes in list of pdf_files and attempts to see if there are findings.
def main():
    pdf_files = ["2021_report.pdf"]
    for pdf_file in pdf_files:
        text_file = pdf_to_txt(pdf_file)
        if flag_findings(text_file):
            print(f"Flagged findings found in {pdf_file}")
        else:
            print(f"No flagged findings found in {pdf_file}")


if __name__ == "__main__":
    main()
