PDF to Excel converter
A comprehensive Streamlit application that provides powerful PDF processing capabilities including PDF to Excel conversion and OCR text extraction from Excel files.
It has 2 tools PDF to Excel Converter and Text Extractor for Excel Images. 

Features
1. PDF to Excel Converter: Convert PDF files to Excel format using Spire.PDF without any format changes.

2. Text Extractor for Excel Images: Extract text from images embedded in Excel files. Make sure to unput the column name in which the extraction to take place. 


Automatically detect and use Tesseract OCR

Support for multiple sheets and column specifications.Flexible text joining options for multiple images in the same cell.Remove images after OCR processing, leaving clean text

**Installation**
Clone the repository:

**bash**
git clone <repository-url>
cd <repository-name>
Install the required dependencies:

**bash**
pip install -r requirements.txt
Install Tesseract OCR:

**Windows:**

Download from Tesseract GitHub releases
Or use Chocolatey: choco install tesseract

**macOS:**

bash
brew install tesseract
Linux (Ubuntu/Debian):

bash
sudo apt-get update
sudo apt-get install tesseract-ocr
Linux (CentOS/RHEL):

bash
sudo yum install tesseract
Usage
Run the main application:

bash
streamlit run app.py
The application will open in your default web browser.Use the sidebar to navigate between different tools:

**PDF to Excel Converter**
Upload a PDF file.Click convert to generate Excel file.Download the converted Excel file

**Text Extractor for Excel Images**
Upload an Excel (.xlsx) file. Configure OCR settings in the sidebar. Tesseract path will be auto-detected if available. Specify target columns for OCR processing. Run OCR to extract text from images

Download the processed Excel file with text instead of images

Project Structure
text
├── app.py                 # Main application launcher
├── PDIG.py              # PDF to Excel converter
├── Text_extractor.py           # Excel Image OCR tool
├── requirements.txt     # Python dependencies
└── README.md           # This file


Dependencies
Key dependencies include:

streamlit - Web application framework
pytesseract - Python wrapper for Tesseract OCR
openpyxl - Excel file manipulation
spire.pdf - PDF processing and conversion
Pillow - Image processing
pymupdf - PDF to image conversion


Maintains tabular structure and formatting where possible

**Excel Image OCR**
Extracts embedded images from Excel cells. Processes images using Tesseract OCR. Replaces images with extracted text. Removes original images from the Excel file. Provides cleaned Excel file for download

**Troubleshooting**
Tesseract Not Found
If Tesseract is not automatically detected:

Ensure Tesseract is installed on your system

Provide the full path to the Tesseract executable in the sidebar

Common paths:

Windows: C:\Program Files\Tesseract-OCR\tesseract.exe
macOS: /usr/local/bin/tesseract
Linux: /usr/bin/tesseract

Conversion Issues
Ensure uploaded files are in supported formats (.pdf for conversion, .xlsx for OCR)

Check file permissions and size limits

Verify all dependencies are properly installed


