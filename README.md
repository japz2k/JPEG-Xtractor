# JPEG Xtractor

**JPEG Xtractor** is a Python-based GUI tool that extracts and copies JPEG files listed in an Excel sheet from a source folder to a structured destination folder.

---

## ✨ Features

- Reads JPEG filenames from an Excel file (`.xlsx` or `.xls`)
- Searches for matching JPEG files in a source directory
- Copies found JPEGs to a destination folder
- Automatically creates a subfolder by replacing "EXCEL" with "PHOTOS" in the Excel filename
- Displays progress with a progress bar and detailed logging
- Simple, user-friendly interface with file and folder pickers

---

## 📦 Installation

### Prerequisites

- Python 3.6 or higher
- Required packages:

```bash
pip install pandas pillow openpyxl tk

```
## 🚀 Running the Application

1. Clone the repository:
   -git clone https://github.com/japz2k/JPEG-Xtractor.git
   -cd JPEG-Xtractor
2. Run the script:
   -python JPEG_Xtractor.py

## 🧭 Usage
Select Excel File
Browse and choose the Excel file containing the list of JPEG filenames (any column is supported).

Select Source Folder
Choose the folder where the original JPEG files are stored.

Select Destination Folder
Choose where to copy the matched JPEGs.

Click "Start"
The tool will scan, match, and copy JPEGs, showing progress and log output.

## 🔁 Example Workflow
Input Excel File: Students_EXCEL.xlsx (JPEG filenames in any column)

Source Folder: C:\Photos\Raw

Destination Folder: C:\Photos\Processed

Output: A new subfolder Students_PHOTOS is created inside Processed, containing all matched JPEGs.
