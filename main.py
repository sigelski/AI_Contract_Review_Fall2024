from nicegui import ui, run, app
import os
import glob
from contract_to_txt import convert_to_txt
from flag_FAR_clauses import annotate_contract
from contract_to_txt import convert_to_txt, txt_to_docx
from flag_problem_language import _flag_problem_language
from nltk import data as nltk

# Append the custom path to the NLTK data path
nltk.path.append("supplementary_files\\nltk_data")

FAR_CLAUSE_MATRIX_PATH = "supplementary_files\\2023-03-20_FAR Matrix.xls"
TNC_MATRIX_PATH = "supplementary_files\\Contract Ts&Cs Matrix.xlsm"

@ui.page("/")
def index():
    ui.add_head_html("<link rel='stylesheet' href='/static/style.css'/>")
    ui.page_title('AI Contract Scanner')

    ui.upload(multiple=True, label="Upload Contracts", auto_upload=True, on_upload=handle_upload).props(add="accept='.docx,.pdf'")


async def handle_upload(e):
    name = e.name
    binary = e.content.read()
    upload_filepath = write_binary_to_temp_file(name, binary)
    downloads_path = os.path.join(os.environ['USERPROFILE'], 'Downloads')
    output_filepath = os.path.join(downloads_path, f"{getFilenameStringNoExtension(name)}_scanned.docx")
    ui.notify("Scanning file...")
    await run.cpu_bound(scan_file, upload_filepath, output_filepath)
    os.remove(upload_filepath)
    ui.notify("Downloading scan")
    ui.download(output_filepath)

def scan_file(filepath, output_filepath):
    # clean docx and annotate it
    convert_to_txt(filepath)
    _flag_problem_language(TNC_MATRIX_PATH)
    back_to_docx = 'flagged_contract_to_txt.txt'
    file_to_highlight = 'flagged_contract_to_docx.docx'
    txt_to_docx(back_to_docx, file_to_highlight)

    annotate_contract(FAR_CLAUSE_MATRIX_PATH, file_to_highlight, output_filepath)

def getFilenameStringNoExtension(filename):
    temp_list = filename.split(".")
    if (len(temp_list) == 1):
        return temp_list[0]
    else:
        temp_list.pop()
        return '.'.join(temp_list)

def write_binary_to_temp_file(name, binary):
    filepath = f"temp/{name}"
    with open(file=filepath, mode="wb") as file:
        file.write(binary)
    return filepath

if __name__ in {"__main__", "__mp_main__"}:
    app.add_static_files("/static", "static")
    app.add_static_files("/temp", "temp")
    ui.run(native=True, reload=False)