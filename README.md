# Automated Insurance Commission Reporting Script

This Python script automates the extraction of commission data from multiple, differently-formatted insurance PDF files. It processes the files, cleans the data, and consolidates everything into a single, clean Excel file ready for analysis.

This script is the data processing engine for a larger automation project that includes an Excel VBA dashboard.

---

### Key Features

*   **Multi-Format PDF Parsing:** Extracts data from 6+ unique carrier PDF layouts.
*   **Data Cleaning & Standardization:** Automatically cleans numbers, dates, and names for consistency.
*   **Automated Excel Output:** Creates a single, tidy `.xlsx` file as the final data source.
*   **Deduplication:** Prevents duplicate entries when running the script multiple times.

### Technologies Used

*   Python 3
*   Pandas
*   Pdfplumber

### How to Use

1.  **Install Libraries:**
    ```bash
    pip install pandas pdfplumber openpyxl
    ```

2.  **Prepare Folders:** Place your PDF files into sub-folders named by the insurance carrier (e.g., `PDFs/CGF/`, `PDFs/Delta/`).

3.  **Run the Script:** Execute the Python script. It will process the PDFs and create/update the final Excel file.
