function main(workbook: ExcelScript.Workbook) {
    // Get the current worksheet.
    let selectedSheet = workbook.getActiveWorksheet();

    // Create a table with the used cells.
    let usedRange = selectedSheet.getUsedRange();
    let newTable = selectedSheet.addTable(usedRange, true);
}