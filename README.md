JPEG Xtractor
A Python-based GUI tool that extracts and copies JPEG files listed in an Excel sheet from a source folder to a destination folder.

Features
Reads JPEG filenames from an Excel file (supports .xlsx and .xls).

Searches for matching JPEGs in a source directory.

Copies found files to a structured destination folder.

Automatically creates a destination folder named after the Excel file (replacing "EXCEL" with "PHOTOS").

Progress bar and logging for tracking operations.

User-friendly interface with file/folder browsing.

Installation
Prerequisites
Python 3.6+

Required packages:

sh
pip install pandas pillow openpyxl tk
Running the Application
Clone the repository:

sh
git clone https://github.com/yourusername/jpeg-xtractor.git
cd jpeg-xtractor
Run the script:

sh
python JPEG_Xtractor.py
Usage
Select Excel File – Browse and choose an Excel file containing JPEG filenames.

Select Source Folder – The directory where the JPEGs are stored.

Select Destination Folder – Where the matched JPEGs will be copied.

Click "Start" – The tool will scan, match, and copy files, displaying progress in the log.

Example Workflow
Input Excel File: Students_EXCEL.xlsx (with JPEG filenames in any column)

Source Folder: C:\Photos\Raw

Destination Folder: C:\Photos\Processed

Output: A new folder Students_PHOTOS is created in Processed, containing matched JPEGs.

Screenshots
(Optional: Add a screenshot of the GUI here.)

License
MIT License

Contributing
Feel free to fork, open issues, or submit pull requests!

Notes for Deployment
For a standalone .exe, use PyInstaller:

sh
pyinstaller --onefile --windowed --icon=assets/jpeg.ico JPEG_Xtractor.py
