function main(workbook: ExcelScript.Workbook, wsKeepName: string,
  visibilityType: string) {

  //Assign worksheet collection to a variable
  let wsArr = workbook.getWorksheets();

  //Create variable to hold visibility type
  let visibilityTypeScript: ExcelScript.SheetVisibility;

  //Loop through all worksheets in the WSArr worksheet collection
  wsArr.forEach(ws => {

    //if the worksheet in the loop is not the worksheet to retain
    if (ws.getName() != wsKeepName) {

      //Use switch statement to select the visibility type
      switch (visibilityType) {
        case "hidden":
          visibilityTypeScript = ExcelScript.SheetVisibility.hidden;

        case "veryHidden":
          visibilityTypeScript = ExcelScript.SheetVisibility.veryHidden;
      }

      //Make the worksheet hidden
      ws.setVisibility(visibilityTypeScript);
    };

  });

}