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

## To run the script by exe file
```bash
pyinstaller --onefile --noconsole --hidden-import="comtypes" --hidden-import="comtypes.client" pdf.py
```

👆This command is to export the python source file to exe file (Take 32MBs of storage)
