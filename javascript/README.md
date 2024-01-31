# Serial Number Matching Script Documentation

This script is designed to match serial numbers from two different worksheets in an Excel workbook. The script operates in two stages to find the best match for each serial number in the "Search" worksheet.

## How it Works

1. The script first retrieves the "Search" and "Desktop" worksheets from the workbook.

2. It then gets the used ranges from both worksheets and extracts their values.

3. The script iterates over each cell in column A of the "Search" worksheet, treating each value as a serial number.

4. For each serial number in the "Search" worksheet, the script checks each cell in column B of the "Desktop" worksheet for a match.

5. If a match is found, the corresponding tag code (from column A of the "Desktop" worksheet) is copied to column B in the "Search" worksheet. The stage of the search (1 in this case) is recorded in column C.

6. If no match is found in the first stage, the script proceeds to the second stage. Here, it looks for the most similar serial number by comparing each character of the serial numbers. However, it only considers those serial numbers whose last characters match with the last character of the search serial number.

7. If the number of matching characters is greater than or equal to the length of the search serial number minus 1, the corresponding tag code is copied to column B in the "Search" worksheet, and the stage of the search (2 in this case) is recorded in column C.

## Usage

This script is particularly useful in scenarios where you have a list of serial numbers in one worksheet ("Search") and you want to find and record the corresponding tag codes from another worksheet ("Desktop"). The two-stage approach helps to find the best match for each serial number, even if an exact match is not found in the first stage.
