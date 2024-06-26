
# Excel File Processing Script
## Overview
This Python script processes multiple Excel files in a specified folder. It reads each file, extracts specific values based on conditions, and saves the results to new Excel files with modified names. Used by Med students in AUTh 

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
