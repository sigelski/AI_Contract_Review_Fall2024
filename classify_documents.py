import os
from contract_to_txt import convert_to_txt
from sklearn.model_selection import train_test_split
from pathlib import Path

# Define directories
TEST_DATA_DIR = Path("supplementary_files/samples_test_set")
CLEAN_DIR = Path("supplementary_files/clean_documents")
FLAGGED_DIR = Path("supplementary_files/flagged_documents")

# Function to check for flagged phrases
def is_flagged(document):
    document_text = convert_to_txt(document)
    if document_text is None:
        # Handle the case where convert_to_txt returns None
        return False
    flagged_phrases = ["POTENTIAL PROBLEMATIC LANGUAGE DETECTED"]
    if "highlight" in document_text:
        return True
    for phrase in flagged_phrases:
        if phrase in document_text:
            return True
    return False

# Get list of documents in test data directory
try:
    documents = os.listdir(TEST_DATA_DIR)
except FileNotFoundError:
    print(f"Error: The directory {TEST_DATA_DIR} does not exist.")
    exit(1)

# Split documents into training and test sets
X_train, X_test, y_train, y_test = train_test_split(documents, documents, test_size=0.1, random_state=42)

# Create directories for clean and flagged documents
CLEAN_DIR.mkdir(exist_ok=True)
FLAGGED_DIR.mkdir(exist_ok=True)

# Move documents to appropriate directories
def classify_document(document):
    try:
        document_path = TEST_DATA_DIR / document
        document_text = convert_to_txt(str(document_path))
        destination_dir = FLAGGED_DIR if is_flagged(document_text) else CLEAN_DIR
        destination_path = destination_dir / document
        os.replace(str(document_path), str(destination_path))
    except FileNotFoundError as fnf_error:
        print(fnf_error)
    except PermissionError as perm_error:
        print(f"Permission error: {perm_error}")
    except OSError as os_error:
        print(f"OS error: {os_error}")
    except Exception as e:
        print(f"An unexpected error occurred while classifying document {document}: {e}")
        
# Classify each document in the test set
for document in X_test:
    classify_document(document)



