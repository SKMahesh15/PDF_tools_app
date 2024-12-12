# PDF Tools Application

### Overview
The PDF Tools Application is a GUI-based tool built using Python and Tkinter that provides a variety of functionalities to manage and manipulate PDF files. This app supports operations like merging PDFs, converting files between PDF and other formats (Word, JPEG), compressing PDFs, and more. It is designed to simplify file management tasks

### Features 
1. Merge PDF Files
Select multiple PDF files and merge them into a single PDF.

2. Word to PDF Conversion
Convert Microsoft Word files (.docx or .doc) into PDF format.

3. JPEG to PDF Conversion
Convert JPEG image files into a PDF.

4. PDF to Word Conversion
Convert PDF files into Microsoft Word format (.docx).

5. PDF to PNG Conversion
Extract pages from a PDF as PNG image files.

### Prerequisites
Ensure the following dependencies and software are installed before running the application:
**Python Libraries:**
`tkinter`
`os`
`pathlib`
`win32com.client` (part of `pywin32` package)
`PyPDF2`
`img2pdf`
`Pillow`
`pdfshrink`
`fitz`(from `PyMuPDF`)

Install the required libraries using pip:
```python 
pip install pywin32 PyPDF2 img2pdf Pillow pdfshrink pymupdf
```

**Microsoft Word Requirement:**
Microsoft Word 2013 or later must be installed for the Word to PDF and PDF to Word functionalities to work properly.

### How to Use
1. Clone or download the repository.
2. Run the application:
```python 
python pdf_tools_app.py
```
3. Use the graphical interface to access the following features:

**Merge PDFs:**
Click on "Merge PDF Files."
Select multiple PDF files and merge them into one.

**Word to PDF:**
Click on "Word to PDF."
Select Word files and convert them to PDF.

**JPEG to PDF:**
Click on "JPEG to PDF."
Select JPEG files and convert them to PDF.

**PDF to Word:**
Click on "PDF to Word."
Select a PDF file and convert it to a Word document.

**PDF to PNG:**
Click on "PDF to PNG."
Select a PDF file and extract its pages as PNG images.

