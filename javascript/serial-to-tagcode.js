function main(workbook: ExcelScript.Workbook) {
  // Get the worksheets
  let searchWorksheet = workbook.getWorksheet("Search");
  let deviceWorksheet = workbook.getWorksheet("Desktop");

  // Get the ranges
  let searchRange = searchWorksheet.getUsedRange();
  let deviceRange = deviceWorksheet.getUsedRange();

  // Get the values
  let searchValues = searchRange.getValues();
  let deviceValues = deviceRange.getValues();

  // Iterate over each cell in column B in the Search worksheet
  for (let i = 0; i < searchValues.length; i++) {
    // Get the serial number from the Search worksheet
    let searchSerialNumber: string = searchValues[i][0];
    let found = false;

    // Iterate over each cell in column B in the worksheet
    for (let j = 0; j < deviceValues.length; j++) {
      // Get the serial number and tagcode from the worksheet
      let deviceSerialNumber: string = deviceValues[j][1];
      let deviceTagCode: string = deviceValues[j][0];

      // If the serial numbers match, copy the tagcode to column B in the Search worksheet
      if (searchSerialNumber === deviceSerialNumber) {
        searchWorksheet.getCell(i, 1).setValue(deviceTagCode);
        found = true;
        break;
      }
    }

    // If not found, search for the last 7 digits
    if (!found) {
      let searchLast7Digits: string = searchSerialNumber.slice(-7);
      for (let j = 0; j < deviceValues.length; j++) {
        let deviceSerialNumber: string = deviceValues[j][1];
        let deviceTagCode: string = deviceValues[j][0];
        let deviceLast7Digits: string = deviceSerialNumber.slice(-7);

        if (searchLast7Digits === deviceLast7Digits) {
          searchWorksheet.getCell(i, 1).setValue(deviceTagCode);
          found = true;
          break;
        }
      }
    }

    // If still not found, find the most similar serial number
    if (!found) {
      let maxMatchCount = 0;
      let maxMatchTagCode = "";
      for (let j = 0; j < deviceValues.length; j++) {
        let deviceSerialNumber: string = deviceValues[j][1];
        let deviceTagCode: string = deviceValues[j][0];
        let matchCount = 0;
        for (let k = 0; k < Math.min(searchSerialNumber.length, deviceSerialNumber.length); k++) {
          if (searchSerialNumber[k] === deviceSerialNumber[k]) {
            matchCount++;
          }
        }
        if (matchCount > maxMatchCount) {
          maxMatchCount = matchCount;
          maxMatchTagCode = deviceTagCode;
        }
      }
      if (maxMatchCount >= searchSerialNumber.length - 1) {
        searchWorksheet.getCell(i, 1).setValue(maxMatchTagCode);
      }
    }
  }
}
