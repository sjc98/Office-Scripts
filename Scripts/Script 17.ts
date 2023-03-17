function main(workbook: ExcelScript.Workbook) {
	let selectedSheet = workbook.getActiveWorksheet();
	// Delete range 1:5 on selectedSheet
	selectedSheet.getRange("1:5").delete(ExcelScript.DeleteShiftDirection.up);
}