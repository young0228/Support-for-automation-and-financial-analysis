Automated Insurance Commission Reporting Pipeline
This repository contains the Python backend for an end-to-end data automation solution designed to streamline financial commission reporting. The script automates the tedious process of extracting data from multiple, differently formatted PDF statements from various insurance carriers, then cleans, standardizes, and consolidates the data for analysis.

This script serves as the data processing engine, feeding a clean dataset into a separate Excel VBA-powered dashboard for interactive visualization and reporting.

🚀 Key Features
Multi-Carrier PDF Parsing: Successfully extracts data from 6+ unique and complex PDF layouts from major providers (CGF, Delta, Jjaramillo, EAIA, Universal, MAPFRE).
Hybrid Extraction Engine:
Utilizes pdfplumber's table extraction for structured, grid-based PDFs.
Employs advanced Regular Expressions (Regex) to parse data from unstructured or text-heavy PDFs where table extraction is not feasible.
Robust Data Cleaning & Standardization: Automatically cleans and formats inconsistent data, including converting numbers in parentheses to negatives, removing currency symbols, and standardizing date formats.
Intelligent Deduplication: Implements logic to merge new data with existing records and removes duplicates based on key identifiers (Policy Number, Period, Insured), ensuring data integrity and preventing double-counting.
Seamless Excel Integration: Outputs a single, clean .xlsx file, which acts as the definitive data source for the Excel VBA frontend, ensuring a clean separation between data processing and presentation.
⚙️ The End-to-End Workflow
This Python script is the first critical step in a larger automated pipeline:

(Multiple PDF Files) ➡️ [insurance_commission.py] ➡️ (Cleaned Data.xlsx) ➡️ [Excel VBA Dashboard] ➡️ (Final Interactive Reports)

🔧 Technologies Used
Python 3
Pandas: For data manipulation, cleaning, and Excel file generation.
Pdfplumber: For robust PDF text and table extraction.
Regular Expressions (Regex): For parsing complex and unstructured text patterns.
Openpyxl: As the engine for writing Pandas DataFrames to .xlsx files.
▶️ Setup and Usage
Clone the Repository

bash
git clone https://github.com/young0228/Support-for-automation-and-financial-analysis.git
cd Support-for-automation-and-financial-analysis
Install Dependencies
It's recommended to use a virtual environment.

bash
pip install -r requirements.txt
(If a requirements.txt file is not present, you can install the libraries manually: pip install pandas pypdf2 pdfplumber openpyxl)

Prepare Folder Structure
The script expects PDF files to be organized in subdirectories named after the carrier within a main PDF folder. For example:

text
/path/to/your/project/
├── 2025/
│   ├── PDF/
│   │   ├── CGF/
│   │   │   └── CGF_Report_2025-01.pdf
│   │   ├── Delta/
│   │   │   └── Delta_Statement_2025-01.pdf
│   │   └── ... (and so on for other carriers)
│   └── 2025.xlsx  (This will be the output file)
└── insurance_commission.py
Configure Paths
Open insurance_commission.py and update the following variables at the bottom of the script to match your folder structure:

python
# Example Configuration
batch_process_data_only(
    '/path/to/your/project/2025/PDF',  # Path to the root PDF folder
    '/path/to/your/project/2025/2025.xlsx', # Path for the output Excel file
)
Run the Script

bash
python insurance_commission.py
The script will process all PDFs, load existing data from the output file (if it exists), merge and deduplicate the records, and save the final clean dataset back to the .xlsx file.

