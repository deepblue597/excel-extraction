
# Excel File Processing Script
## Overview

This project automates the extraction of specific medical test values from Excel files, for medical students at AUTh. It identifies key biomarkers (e.g., glucose, urea, creatinine, SGOT) from structured lab reports, processes the data, and saves the extracted values in a new Excel file for further analysis.

## Requirements
- Python 3.x
-  pandas library
- glob library (part of the standard library)
- os library (part of the standard library)

## Setup
Ensure you have Python 3.x installed on your system.
Install the pandas library if you haven't already:
`pip install pandas`
## Usage
1. Modify the Folder Path:
Replace 'C:/test_excel' with the actual path to your folder containing the Excel files:
`folder_path = 'C:/your_folder_path'`

2. Place Your Excel Files:
Place the Excel files with `.xlsx` and `.xlt` extensions in the specified folder.

3. Run the Script:
Execute the script to process the Excel files. The script will:

- Read each Excel file.-
-  Extract specific values from the files based on given conditions.
-  Save the extracted values into new Excel files with modified names.
