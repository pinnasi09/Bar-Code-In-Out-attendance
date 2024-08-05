# Bar-Code-In-Out-attendance
## Description 
this model is made for "LIBRARY " Maintaining log book  is difficult this is why i have created this model .
# Barcode Scanning System

This project implements a barcode scanning system using Python, OpenCV, and SQLite. The system captures barcode data via a webcam, logs the scans into a SQLite database, and generates an Excel report of the recorded data.

## Features

- **Barcode Scanning**: Continuously scans barcodes using a webcam.
- **Database Logging**: Stores scan data in a SQLite database.
- **Automatic Beep**: Emits a beep sound on successful scan (Windows only).
- **Excel Export**: Generates an Excel file containing all scan records.
- **Automatic Time Management**: Tracks both in and out times for each barcode.

## Requirements

- Python 3.x
- `opencv-python`
- `pyzbar`
- `sqlite3` (part of Python standard library)
- `openpyxl`
- `winsound` (Windows only)

## Installation

1. **Clone the repository:**

   ```sh
   git clone https://github.com/yourusername/barcode-scanning-system.git
   cd barcode-scanning-system
Install the required packages:

## sh

pip install opencv-python pyzbar openpyxl

Note: winsound is available by default on Windows. For other platforms, you may need to adjust or remove beep functionality.

Usage
Create the database and start the barcode scanner:

Run the script:

sh

python barcode_scanner.py

## Barcode Scanning:

The script will open your webcam and start scanning for barcodes.
Detected barcodes will be logged into the SQLite database.
A beep sound will be emitted each time a barcode is detected (Windows only).
Exit the Program:

Press q while the scanning window is active to exit the program.
Data Export:

After exiting, an Excel file containing the scan data will be generated in the project directory.
File Structure
barcode_scanner.py: The main script that handles barcode scanning, database operations, and Excel export.
scans_YYYYMMDD_HHMMSS.xlsx: Generated Excel file with scan data.
Notes
Ensure that your webcam is properly connected and accessible.
The script uses winsound for beep sound which is available on Windows. For other platforms, consider modifying the beep functionality.
The database and Excel file names include timestamps to avoid overwriting previous files.
License
This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments
 OpenCV for computer vision capabilities.
 pyzbar for barcode decoding.
**SQLite for lightweight database storage.
openpyxl for Excel file handling.**
Feel free to customize the repository URL and license information as needed!


Feel free to modify any details, such as the repository URL or specific acknowledgments, to suit your project and preferences.



