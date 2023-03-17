function main(workbook: ExcelScript.Workbook) {
    // Get the current worksheet.
    let selectedSheet = workbook.getActiveWorksheet();

    // Create a table with the used cells.
    let usedRange = selectedSheet.getUsedRange();
    let newTable = selectedSheet.addTable(usedRange, true);

    // Sort the table using the first column.
    newTable.getSort().apply([{ key: 0, ascending: true }]);
}