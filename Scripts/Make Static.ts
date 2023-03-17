function main(workbook: ExcelScript.Workbook) {
    // Get the ranges with formulas.
    let sheet = workbook.getActiveWorksheet();
    let usedRange = sheet.getUsedRange();
    let formulaCells = usedRange.getSpecialCells(ExcelScript.SpecialCellType.formulas);

    // In each formula range: get the current value, clear the contents, and set the value as the old one.
    // This removes the formula but keeps the result.
    formulaCells.getAreas().forEach((range) => {
        let currentValues = range.getValues();
        range.clear(ExcelScript.ClearApplyTo.contents);
        range.setValues(currentValues);
    });
}