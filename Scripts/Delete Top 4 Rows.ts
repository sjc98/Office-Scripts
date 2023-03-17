function main(workbook: ExcelScript.Workbook) {
	let selectedSheet = workbook.getActiveWorksheet();
	// Delete range 1:4 on selectedSheet
	selectedSheet.getRange("1:4").delete(ExcelScript.DeleteShiftDirection.up);
}