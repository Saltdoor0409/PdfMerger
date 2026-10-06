# PdfMerger
PDF merger Support Word, PPT/PPTX files convert to PDF and merge together

# PDF Merger & Converter Installation Guide

## 1. Prerequisites
- **OS**: Windows only.
- **Software**: Local installation of Microsoft Office (Word & PowerPoint).
- **Environment**: Python 3.x installed.

## 2. Install Dependencies
Open Command Prompt (CMD) and execute the following command:
```bash
pip install pypdf comtypes
```

## To run the script directly from the source code
```bash
python wordpptpdf.py
```

## To run the script by exe file (Optional)
```bash
pyinstaller --onefile --noconsole --hidden-import="comtypes" --hidden-import="comtypes.client" pdf.py
```

👆This command is to export the python source file to exe file (Take 32MBs of storage)

## Features

- **Multi-Format Support**
  Add PDF, Word (`.doc`, `.docx`), and PowerPoint (`.ppt`, `.pptx`) files directly. 

- **Automated Conversion**
  Automatically converts Word and PowerPoint documents to PDF in the background using the local Microsoft Office COM interface. 

- **Flexible Sorting**
  Reorder files manually (Move Up / Move Down), or auto-sort them by File Name or Modified Date in Ascending/Descending order. 

- **List Management**
  Remove specific selected files or clear the entire list with a single click. 

- **One-Click Merge**
  Combine all processed files in the current list and save them as a single, brand-new PDF document.
