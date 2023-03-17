function main(workbook: ExcelScript.Workbook) {
    // Get the current worksheet.
    let selectedSheet = workbook.getActiveWorksheet();

    // Get the value of cell A1.
    let range = selectedSheet.getRange("A1");

    // Print the value of A1.
    console.log(range.getValue());
}