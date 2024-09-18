# Sheet Analyzing Project - README

## Overview

This project provides a Python script designed to compare and analyze data from two Excel files using fuzzy string matching. The script takes data from two files, Mark Meyers.xlsx and Shauns Report.xlsx, and performs fuzzy matching on specified columns. The matching results are filtered and saved into a new Excel file for further analysis.

## Features

- Upload and process two Excel files.
- Extract relevant data from multiple sheets.
- Perform fuzzy string matching using the `fuzzywuzzy` library on specific columns.
- Filter and save matched data to a new Excel file.
- The tool is customizable to adjust matching thresholds and column names for comparison.

## Dependencies

The script requires the following Python libraries to be installed:

- `pandas`: For reading and manipulating Excel data.
- `fuzzywuzzy[speedup]`: For performing fuzzy string matching.
- `openpyxl`: For working with Excel files.
- `google.colab`: Required only if running in Google Colab for file uploads.

To install the required libraries, use the following command:
```bash
pip install fuzzywuzzy[speedup] openpyxl pandas
```

## Files

- Mark Meyers.xlsx: One of the input Excel files with multiple sheets containing company-related data.
- Shauns Report.xlsx: Another input Excel file to be compared with data from the **Mark Meyers.xlsx** file.
- Filtered_Shauns_Report.xlsx: The resulting file containing the filtered data after fuzzy matching.

## How the Code Works

1. Upload Files: 
   - The user is prompted to upload Mark Meyers.xlsx and Shauns Report.xlsx files.
   
2. Process Mark Meyers File:
   - The code reads all sheets from the **Mark Meyers.xlsx** file.
   - It extracts unique rows from the `Product Category` column (if available) and lists the sheet names.

3. **Fuzzy Matching**:
   - Three columns are selected for comparison:
     - `By_Bill_To_Company`: Bill-to-company column.
     - `By_Ship_To_Company`: Ship-to-company column.
     - `FBO Name` in **Shauns Report.xlsx**.
   - The `fuzzywuzzy` library is used to perform fuzzy string matching on these columns with a similarity threshold set at 85 (adjustable).

4. **Filter and Save Data**:
   - The code filters **Shauns Report.xlsx** to retain only the rows that have matches above the threshold.
   - The matched data is saved as **Filtered_Shauns_Report.xlsx**.

## How to Run the Code

### Google Colab Instructions

1. Open the notebook in Google Colab.
2. Install the necessary libraries by running the provided `pip` command.
3. Upload both **Mark Meyers.xlsx** and **Shauns Report.xlsx** files when prompted.
4. The script will automatically perform the comparison and save the filtered results as **Filtered_Shauns_Report.xlsx**.

### Local Machine Instructions

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/sheet_analyzing.git 
   ```
2. Install the dependencies:
   ```bash
   pip install fuzzywuzzy[speedup] openpyxl pandas
   ```
3. Place the input Excel files (`Mark Meyers.xlsx` and `Shauns Report.xlsx`) in the same directory as the script.
4. Run the script:
   ```bash
   python sheet_analyzing.py
   ```
5. The filtered report will be saved as **Filtered_Shauns_Report.xlsx**.

## Customization

- **Threshold**: You can adjust the fuzzy matching threshold in the `threshold` variable.
- **Column Selection**: Modify the column names used for comparison in the `fuzzy_match` function calls.
  

## Author

Deep Patel

## Acknowledgments

- This project uses the `fuzzywuzzy` library for fuzzy matching.

