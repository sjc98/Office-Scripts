function main(workbook: ExcelScript.Workbook) {
    // Get the current active cell in the workbook.
    let cell = workbook.getActiveCell();

    // Log that cell's value.
    console.log(`The current cell's value is ${cell.getValue()}`);
}