JPEG Xtractor

A Python-based GUI tool that extracts and copies JPEG files listed in an Excel sheet from a source folder to a destination folder.

Features
  •	Reads JPEG filenames from an Excel file (supports .xlsx and .xls).
  •	Searches for matching JPEGs in a source directory.
  •	Copies found files to a structured destination folder.
  •	Automatically creates a destination folder named after the Excel file (replacing "EXCEL" with "PHOTOS").
  •	Progress bar and logging for tracking operations.
  •	User-friendly interface with file/folder browsing.

Installation

Prerequisites
  •	Python 3.6+
  • Required package
      -- pip install pandas pillow openpyxl tk
      
Running the Application
1.	Clone the repository:
  •	git clone https://github.com/japz2k/JPEG-Xtractor.git
  •	cd jpeg-xtractor
2.	Run the script:
  •	python JPEG_Xtractor.py

Usage
  1.	Select Excel File – Browse and choose an Excel file containing JPEG filenames.
  2.	Select Source Folder – The directory where the JPEGs are stored.
  3.	Select Destination Folder – Where the matched JPEGs will be copied.
  4.	Click "Start" – The tool will scan, match, and copy files, displaying progress in the log.
Example Workflow
  •	Input Excel File: Students_EXCEL.xlsx (with JPEG filenames in any column)
  •	Source Folder: C:\Photos\Raw
  •	Destination Folder: C:\Photos\Processed
  •	Output: A new folder PHOTOS is created in Processed, containing matched JPEGs.

License
MIT License

Contributing
Feel free to fork, open issues, or submit pull requests!
________________________________________
