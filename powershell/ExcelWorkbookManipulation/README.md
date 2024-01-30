# PowerShell Script for Excel Workbook Manipulation

This PowerShell script is designed to manipulate an Excel workbook. It opens a specified workbook, checks if it's read-only, and if so, changes its attribute to normal. It then iterates over the worksheets and performs operations on the used range of each worksheet.

## Features

- Opens a specified Excel workbook.
- Checks if the workbook is read-only. If it is, it changes the workbook's attribute to normal.
- Retrieves the number of worksheets in the workbook.
- Iterates over each worksheet and performs operations on the used range of each worksheet.
- For each row in the used range, it retrieves the values of certain columns and performs operations based on these values.
- Updates a specific column with a specified value.
- Saves and closes the workbook.

## Usage

1. Replace the `$filePath` variable with the path to the Excel workbook you want to manipulate.
2. Run the script in a PowerShell environment.

## Requirements

- Excel must be installed on the machine where the script is run.
- The user running the script must have the necessary permissions to change the attributes of the file and to perform operations on the Excel workbook.

## Note

This script is intended to be used with workbooks that have a specific structure. If your workbook is structured differently, you may need to adjust the script accordingly.

***Please remember to handle all data with care and ensure you have the necessary permissions before running the script. 😊***
