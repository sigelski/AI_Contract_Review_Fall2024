from flag_FAR_clauses import annotate_contract
import tkinter
from tkinter import filedialog
from contract_to_txt import convert_to_txt, txt_to_docx
from flag_problem_language import _flag_problem_language
from classify_documents import classify_document
from nltk import data as nltk
from pathlib import Path

# Append the custom path to the NLTK data path
nltk.path.append("supplementary_files\\nltk_data")

tkinter.Tk().withdraw()
# Opens file browser to select Excel spreadsheet containing FAR Clauses
print("Loading FAR Matrix")
FAR_clause_matrix = Path("supplementary_files/_FAR_Matrix.xls")

# Opens file browser to select Excel spreadsheet containing FAR Clauses
print("Loading AU T&Cs Matrix")
tnc_matrix = Path("supplementary_files/_Contract_Ts&Cs_Matrix.xlsm")

# Opens file browser to select document you wish to parse and annotate according to selected spreadsheet
print("Select contract you wish to annotate")
contract_in = filedialog.askopenfilename()

# Opens file browser and prompts user to save a file and input name for file
print("Select the directory you wish to save annotated contract to")
save_path = filedialog.asksaveasfilename()

# clean docx and annotate it
convert_to_txt(contract_in)
_flag_problem_language(tnc_matrix)
back_to_docx = 'flagged_contract_to_txt.txt'
file_to_highlight = 'flagged_contract_to_docx.docx'
txt_to_docx(back_to_docx, file_to_highlight)

annotate_contract(FAR_clause_matrix, file_to_highlight, save_path)

print("Annotation complete. Output file: " + save_path)

classify_document(contract_in)
