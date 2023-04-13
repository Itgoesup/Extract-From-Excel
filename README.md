
# Excel Data Extraction Script


This is a simple Python script that extracts data from a specific sheet in an Excel workbook and creates a new Excel workbook with the extracted data.








## Installation

To run this script, you will need the openpyxl module. You can install it by running:

```bash
pip install openpyxl

```
    
## Usage

1. Place the Combined To-Extract.xlsx file in the same directory as the script.
2. Specify your parameters by modifying the following variables in the script:
 - Sheet_name: the name of the sheet in the source Excel file that you want to extract data from.
- Which_to_find: the specific string or word you want to extract data for.
- New_Sheet_Name: the name of the new Excel file that will be created with the extracted data.

3. Run the script by executing the ipynb

The extracted data will be saved in a new Excel file with the specified name in the same directory as the script.

## Note

Please make sure to close all Excel files that you are about to access before running the script.
