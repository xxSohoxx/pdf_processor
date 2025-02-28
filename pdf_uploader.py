import os
import shutil
import magic
import re
import sys

from pdf_processor import process_pdf_folder, save_to_excel

from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QListWidget, QTextEdit

UPLOAD_FOLDER = "uploads"

def clear_upload_folder():
    """Delete all files in the upload folder before starting."""
    if os.path.exists(UPLOAD_FOLDER):
        shutil.rmtree(UPLOAD_FOLDER)  # Delete folder and all contents
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)  # Recreate an empty folder

clear_upload_folder()  # Ensure the folder is empty before execution

def allowed_file(filepath):
    """Check file name and MIME type."""
    filename = os.path.basename(filepath)
    if not re.fullmatch(r"[a-zA-Z0-9_-]+\.pdf", filename):
        return False, "Invalid file name"

    mime = magic.Magic(mime=True)
    file_mime_type = mime.from_file(filepath)

    return file_mime_type == "application/pdf", f"Invalid MIME type: {file_mime_type}"

class PDFUploader(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("PDF Uploader & Processor")
        self.setGeometry(100, 100, 600, 400)

        layout = QVBoxLayout()

        self.label = QLabel("Upload PDF Files")
        layout.addWidget(self.label)

        self.upload_button = QPushButton("Select Files")
        self.upload_button.clicked.connect(self.upload_files)
        layout.addWidget(self.upload_button)

        self.file_list = QListWidget()
        layout.addWidget(self.file_list)

        self.process_button = QPushButton("Process Files")
        self.process_button.clicked.connect(self.process_files)
        layout.addWidget(self.process_button)

        self.output_log = QTextEdit()
        self.output_log.setReadOnly(True)
        layout.addWidget(self.output_log)

        self.setLayout(layout)

    def upload_files(self):
        """Select, validate, and upload files."""
        filepaths, _ = QFileDialog.getOpenFileNames(self, "Select PDF files", "", "PDF Files (*.pdf)")
        if not filepaths:
            return
        
        self.file_list.clear()

        shutil.rmtree(UPLOAD_FOLDER, ignore_errors=True)
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)

        for filepath in filepaths:
            is_valid, reason = allowed_file(filepath)
            filename = os.path.basename(filepath)

            if is_valid:
                destination = os.path.join(UPLOAD_FOLDER, filename)
                shutil.copy2(filepath, destination)  # COPY instead of MOVE
                self.file_list.addItem(f"✅ {filename} copied successfully.")
            else:
                self.file_list.addItem(f"❌ {filename} rejected: {reason}")

    def process_files(self):
        """Process PDFs without threading (to check for errors)."""
        self.output_log.clear()
        self.output_log.append("🔄 Processing...")  
        QApplication.processEvents()  
    
        extracted_data = process_pdf_folder()
        if extracted_data:
            save_to_excel(extracted_data)
            self.output_log.append("✅ Processing completed. Data saved to Excel.")
        else:
            self.output_log.append("⚠ No valid PDFs found in uploads.")

# Run application
app = QApplication(sys.argv)
window = PDFUploader()
window.show()
sys.exit(app.exec())
