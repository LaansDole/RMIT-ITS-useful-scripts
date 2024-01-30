# Search Tagcode from Serialnumber
This script is designed to search for a serial number in a given Excel workbook. It operates on two worksheets: “Search” and “Desktop”.

## Functionality
**The script performs the following steps:**

* Retrieves the “Search” and “Desktop” worksheets from the workbook.
* Gets the used ranges from both worksheets.
* Retrieves the values from these ranges.
* Iterates over each cell in column B of the “Search” worksheet.
* For each cell, it retrieves the serial number and checks if it matches any serial number in the “Desktop” worksheet.
* If a match is found, it copies the corresponding tag code from the “Desktop” worksheet to column B of the “Search” worksheet.
* If no match is found, it then checks if the last 7 digits of the serial number match any serial number in the “Desktop” worksheet.
* If a match is found, it copies the corresponding tag code from the “Desktop” worksheet to column B of the “Search” worksheet.
* If still no match is found, it finds the most similar serial number in the “Desktop” worksheet (i.e., the serial number that matches at least 6 characters with the search serial number).
* If a similar serial number is found, it copies the corresponding tag code from the “Desktop” worksheet to column B of the “Search” worksheet.
## Usage
This script is intended to be used in the Office Scripts for Excel environment. To use the script, simply open the Excel workbook of interest and run the script. The script will automatically perform the serial number search and update the “Search” worksheet with the corresponding tag codes.

***Please note that the script assumes the serial numbers are located in column B of both worksheets and the tag codes are located in column A of the “Desktop” worksheet. If your data is arranged differently, you may need to adjust the script accordingly. 😊***
