function main(workbook: ExcelScript.Workbook, worksheetToKeepName: string,
  visibilityType: string) {


  let worksheets = workbook.getWorksheets();

  //Unprotect all worksheets in workbook - Method #1
  worksheets.forEach(ws => {
    ws.getProtection().unprotect();
  });

  // Unhide all worksheets
  workbook.getWorksheets().forEach(ws => {
    ws.setVisibility(ExcelScript.SheetVisibility.visible);
  });

  // Get the active worksheet
  let sheet = workbook.getActiveWorksheet();

  // Get the used range in the active worksheet
  let usedRange = sheet.getUsedRange();

  // Assign worksheet collection to a variable
  let wsArr = workbook.getWorksheets();

  // Create variable to hold visibility type
  let visibilityTypeScript: ExcelScript.SheetVisibility;

  // Loop through all worksheets in the WSArr worksheet collection
  wsArr.forEach(ws => {

    // If the worksheet in the loop is not the worksheet to retain
    if (ws.getName() != worksheetToKeepName) {

      // Use switch statement to select the visibility type
      switch (visibilityType) {
        case "hidden":
          visibilityTypeScript = ExcelScript.SheetVisibility.hidden;
          break;
        case "veryHidden":
          visibilityTypeScript = ExcelScript.SheetVisibility.veryHidden;
          break;
        default:
          console.log("Error: Invalid visibility type. Defaulting to 'hidden'.")
          visibilityTypeScript = ExcelScript.SheetVisibility.hidden;
      }

    };

  });

}