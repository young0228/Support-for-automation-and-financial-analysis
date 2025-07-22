Automated Insurance Commission Reporting Pipeline
This project is the Python component of an end-to-end data pipeline designed to automate the process of financial commission reporting. It extracts data from multiple, differently formatted PDF statements from various insurance carriers, cleans and standardizes the data, and prepares it for analysis in an Excel VBA-powered dashboard.

Key Features
Multi-Format PDF Parsing: Extracts commission data from 6+ unique PDF layouts (from providers like CGF, Delta, Universal, MAPFRE, etc.).
Advanced Data Extraction: Uses pdfplumber for table extraction and advanced Regular Expressions (Regex) for text-based and semi-structured data.
Data Cleaning & Standardization: Cleans numerical values, standardizes dates, and structures the output for consistency.
Deduplication Logic: Identifies and removes duplicate records when merging new data with existing records.
Excel Integration: Outputs a clean .xlsx file, which serves as the data source for a corresponding Excel VBA dashboard.
Technologies Used
Python 3
Libraries: pandas, pypdf2, pdfplumber, openpyxl
How to Use
Install the required libraries: pip install -r requirements.txt
Organize your PDF files into subdirectories within a root folder (e.g., /PDFs/CGF/, /PDFs/Delta/).
Update the root_folder and data_file_path variables in the script to match your environment.
Run the script: python insurance_commission.py
The consolidated and cleaned data will be saved to the specified Excel file.
Note: This script is the data processing engine. The final analysis and visualization are handled by a separate Excel VBA module.
