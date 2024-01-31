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

  // Iterate over each cell in column A in the Search worksheet
  for (let i = 0; i < searchValues.length; i++) {
    // Get the serial number from the Search worksheet
    let searchSerialNumber: string = searchValues[i][0].toString().toUpperCase();
    let found = false;
    let stage = 1;

    // Iterate over each cell in column B in the device worksheet
    for (let j = 0; j < deviceValues.length; j++) {
      // Get the serial number and tagcode from the worksheet
      let deviceSerialNumber: string = deviceValues[j][1].toString().toUpperCase();
      let deviceTagCode: string = deviceValues[j][0];

      // If the serial numbers match, copy the tagcode to column B in the Search worksheet
      if (searchSerialNumber === deviceSerialNumber) {
        searchWorksheet.getCell(i, 1).setValue(deviceTagCode);
        searchWorksheet.getCell(i, 2).setValue("Stage " + stage);
        found = true;
        break;
      }
    }

    // If still not found, find the most similar serial number
    if (!found) {
      stage++;
      let maxMatchCount = 0;
      let maxMatchTagCode = "";
      for (let j = 0; j < deviceValues.length; j++) {
        let deviceSerialNumber: string = deviceValues[j][1].toString().toUpperCase();
        let deviceTagCode: string = deviceValues[j][0];
        let matchCount = 0;
        if (searchSerialNumber[searchSerialNumber.length - 1] === deviceSerialNumber[deviceSerialNumber.length - 1]) {
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
      }
      if (maxMatchCount >= searchSerialNumber.length - 1) {
        searchWorksheet.getCell(i, 1).setValue(maxMatchTagCode);
        searchWorksheet.getCell(i, 2).setValue("Stage " + stage);
      }
    }
  }
}
