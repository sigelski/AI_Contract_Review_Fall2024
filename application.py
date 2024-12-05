import sys
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton,
    QTextEdit, QVBoxLayout, QHBoxLayout, QFileDialog,
    QMessageBox, QProgressBar
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon, QPixmap, QFont
import qtawesome as qta

# Custom Modules
from contract_to_txt import convert_to_txt, txt_to_docx
from flag_FAR_clauses import annotate_contract
from flag_problem_language import flag_problem_language
from docx import Document

import nltk
script_dir = os.path.dirname(os.path.abspath(__file__))
nltk_data_dir = os.path.join(script_dir, 'supplementary_files', 'nltk_data')
nltk.data.path.append(nltk_data_dir)

# Defines paths to supplementary files
FAR_CLAUSE_MATRIX_PATH = os.path.join('supplementary_files', '2023-03-20_FAR Matrix.xls')
TNC_MATRIX_PATH = os.path.join('supplementary_files', 'Contract Ts&Cs Matrix.xlsm')

class Worker(QThread):
    progress_update = pyqtSignal(int)
    analysis_complete = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self, contract_in, output_filepath):
        super().__init__()
        self.contract_in = contract_in
        self.output_filepath = output_filepath

    def run(self):
        try:
            # Update progress: Start
            self.progress_update.emit(10)

            # Convert to text
            convert_to_txt(self.contract_in)
            self.progress_update.emit(30)

            # Flag problem language
            flag_problem_language(TNC_MATRIX_PATH)
            self.progress_update.emit(50)

            # Convert back to docx
            back_to_docx = 'flagged_contract_to_txt.txt'
            file_to_highlight = 'flagged_contract_to_docx.docx'
            txt_to_docx(back_to_docx, file_to_highlight)
            self.progress_update.emit(70)

            # Annotate contract
            annotate_contract(FAR_CLAUSE_MATRIX_PATH, file_to_highlight, self.output_filepath)
            self.progress_update.emit(90)

            # Finalize
            self.progress_update.emit(100)

            # Signal that analysis is complete
            self.analysis_complete.emit(self.output_filepath)
        except Exception as e:
            import traceback
            error_message = f"An error occurred during analysis:\n{traceback.format_exc()}"
            self.error_occurred.emit(error_message)

class ContractReviewApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('AI Contract Review Tool')
        self.resize(800, 600)
        self.initUI()

    def initUI(self):
        # Set the overall font
        self.setFont(QFont('Arial', 10))

        # Set the main window background color
        self.setStyleSheet("background-color: #F5F5F5;")

        # Layouts
        self.layout = QVBoxLayout()

        # Auburn University logo (ensure the path is correct)
        logo_path = os.path.join('supplementary_files', 'logos', 'au_logo_fullcolor_blue.png')
        if os.path.exists(logo_path):
            self.logo_label = QLabel()
            self.logo_label.setPixmap(QPixmap(logo_path).scaled(150, 150, Qt.KeepAspectRatio))
            self.logo_label.setAlignment(Qt.AlignCenter)
            self.layout.addWidget(self.logo_label)
        else:
            # If logo not found, display a placeholder
            self.logo_label = QLabel('Auburn University Logo')
            self.logo_label.setAlignment(Qt.AlignCenter)
            self.logo_label.setStyleSheet("color: #0C2340; font-weight: bold;")
            self.layout.addWidget(self.logo_label)

        # Title Label
        title_font = QFont('Arial', 16, QFont.Bold)
        self.title_label = QLabel('AI Contract Review Tool')
        self.title_label.setFont(title_font)
        self.title_label.setAlignment(Qt.AlignCenter)
        self.title_label.setStyleSheet("color: #0C2340;")  # Navy Blue
        self.layout.addWidget(self.title_label)

        # Widgets
        self.upload_btn = QPushButton(' Upload Contract')
        self.upload_btn.clicked.connect(self.open_file_dialog)
        self.upload_btn.setMinimumHeight(40)
        self.upload_btn.setIcon(qta.icon('fa.folder-open', color='white'))
        self.upload_btn.setStyleSheet("""
            QPushButton {
                background-color: #0C2340;  /* Navy Blue */
                color: white;
                font-weight: bold;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #E87722;  /* Burnt Orange */
            }
        """)

        self.analyze_btn = QPushButton(' Analyze Contract')
        self.analyze_btn.clicked.connect(self.analyze_contract)
        self.analyze_btn.setMinimumHeight(40)
        self.analyze_btn.setIcon(qta.icon('fa.play', color='white'))
        self.analyze_btn.setStyleSheet("""
            QPushButton {
                background-color: #0C2340;
                color: white;
                font-weight: bold;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #E87722;
            }
        """)

        self.open_output_btn = QPushButton(' Open Analyzed Contract')
        self.open_output_btn.clicked.connect(self.open_output_file)
        self.open_output_btn.setEnabled(False)
        self.open_output_btn.setMinimumHeight(40)
        self.open_output_btn.setIcon(qta.icon('fa.file-word-o', color='white'))
        self.open_output_btn.setStyleSheet("""
            QPushButton {
                background-color: #0C2340;
                color: white;
                font-weight: bold;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #E87722;
            }
        """)

        # Horizontal layout for buttons
        button_layout = QHBoxLayout()
        button_layout.addWidget(self.upload_btn)
        button_layout.addWidget(self.analyze_btn)
        button_layout.addWidget(self.open_output_btn)

        self.contract_text = QTextEdit()
        self.contract_text.setReadOnly(True)
        self.contract_text.setPlaceholderText('Contract content will appear here.')
        self.contract_text.setStyleSheet("""
            background-color: white;
            border: 2px solid #0C2340;
            border-radius: 5px;
            padding: 10px;
        """)

        self.message_label = QLabel('')
        self.message_label.setWordWrap(True)
        self.message_label.setStyleSheet("color: #0C2340; font-weight: bold;")
        self.message_label.setAlignment(Qt.AlignCenter)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #0C2340;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #E87722;
                width: 20px;
            }
        """)

        # Add widgets to layout
        self.layout.addLayout(button_layout)
        self.layout.addWidget(self.contract_text)
        self.layout.addWidget(self.message_label)
        self.layout.addWidget(self.progress_bar)

        self.layout.setSpacing(15)
        self.layout.setContentsMargins(20, 20, 20, 20)

        self.setLayout(self.layout)

    def open_file_dialog(self):
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select Contract File",
            "",
            "Word Documents (*.docx);;PDF Files (*.pdf);;Text Files (*.txt);;All Files (*)",
            options=options
        )
        if filename:
            try:
                extension = os.path.splitext(filename)[1].lower()
                if extension == '.docx':
                    # Handle .docx files
                    doc = Document(filename)
                    content = '\n'.join([para.text for para in doc.paragraphs])
                elif extension == '.pdf':
                    # Handle .pdf files
                    pdf_text = self.read_pdf(filename)
                    content = pdf_text
                else:
                    # Specify the encoding explicitly
                    with open(filename, 'r', encoding='utf-8') as file:
                        content = file.read()
                self.contract_text.setPlainText(content)
                self.contract_path = filename  # Store the file path for later use

                # Clear previous messages
                self.message_label.setText('')
                self.open_output_btn.setEnabled(False)  # Disable the open button

            except UnicodeDecodeError as e:
                # Try a different encoding or inform the user
                self.contract_text.setPlainText(f"Failed to decode the file using UTF-8 encoding: {e}")
            except Exception as e:
                self.contract_text.setPlainText(f"Failed to load file: {e}")

    def read_pdf(self, pdf_path):
        try:
            pdf_reader = PdfReader(pdf_path)
            pdf_text = ''
            for page in pdf_reader.pages:
                pdf_text += page.extract_text()
            return pdf_text
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to read PDF file: {e}")
            return ""

    def analyze_contract(self):
        if hasattr(self, 'contract_path') and self.contract_path:
            try:
                # Prompt user to select save location
                options = QFileDialog.Options()
                options |= QFileDialog.DontUseNativeDialog
                default_save_name = os.path.splitext(os.path.basename(self.contract_path))[0] + '_scanned.docx'
                save_filepath, _ = QFileDialog.getSaveFileName(
                    self,
                    "Save Analyzed Contract As",
                    os.path.join(os.path.dirname(self.contract_path), default_save_name),
                    "Word Documents (*.docx);;All Files (*)",
                    options=options
                )

                if not save_filepath:
                    # User canceled the save dialog
                    self.message_label.setText("Analysis canceled. Please select a save location to proceed.")
                    return

                # Initialize the progress bar
                self.progress_bar.setValue(0)
                self.progress_bar.setVisible(True)

                # Disable buttons to prevent multiple clicks
                self.upload_btn.setEnabled(False)
                self.analyze_btn.setEnabled(False)
                self.open_output_btn.setEnabled(False)

                # Clear previous messages
                self.message_label.setText('')

                # Create and start the worker thread
                self.worker = Worker(self.contract_path, save_filepath)
                self.worker.progress_update.connect(self.update_progress_bar)
                self.worker.analysis_complete.connect(self.analysis_finished)
                self.worker.error_occurred.connect(self.analysis_error)
                self.worker.start()

            except Exception as e:
                import traceback
                error_message = f"An error occurred during analysis:\n{traceback.format_exc()}"
                self.message_label.setText(error_message)
        else:
            self.message_label.setText("Please upload a contract first.")

    def update_progress_bar(self, value):
        self.progress_bar.setValue(value)

    def analysis_finished(self, output_filepath):
        # Hide the progress bar
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)

        # Re-enable buttons
        self.upload_btn.setEnabled(True)
        self.analyze_btn.setEnabled(True)

        # Display the completion message
        self.message_label.setText(f"Analysis complete. Output file saved as:\n{output_filepath}")

        # Store the output file path
        self.output_filepath = output_filepath

        # Enable the button to open the output file
        self.open_output_btn.setEnabled(True)

    def analysis_error(self, error_message):
        # Display the error in the message label
        self.message_label.setText(error_message)

        # Hide the progress bar
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)

        # Re-enable buttons
        self.upload_btn.setEnabled(True)
        self.analyze_btn.setEnabled(True)

    def open_output_file(self):
        if hasattr(self, 'output_filepath') and os.path.exists(self.output_filepath):
            try:
                if sys.platform.startswith('darwin'):
                    # macOS
                    import subprocess
                    subprocess.call(('open', self.output_filepath))
                elif os.name == 'nt':
                    # Windows
                    os.startfile(self.output_filepath)
                elif os.name == 'posix':
                    # Linux
                    import subprocess
                    subprocess.call(('xdg-open', self.output_filepath))
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to open the file: {e}")
        else:
            QMessageBox.warning(self, "File Not Found", "The output file does not exist.")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ContractReviewApp()
    ex.show()
    sys.exit(app.exec_())