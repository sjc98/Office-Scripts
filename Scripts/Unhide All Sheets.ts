/**
 * This script unhides all the worksheets in the workbook.
 */
function main(workbook: ExcelScript.Workbook) {
  // Iterate over each worksheet.
  workbook.getWorksheets().forEach((worksheet) => {
    // Set the worksheet visibility to visible.
    worksheet.setVisibility(ExcelScript.SheetVisibility.visible);
  });
}