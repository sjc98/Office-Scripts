function main(workbook: ExcelScript.Workbook) {
	let selectedSheet = workbook.getActiveWorksheet();
	// Auto fit the columns of range range A:H on selectedSheet
	selectedSheet.getRange("A:H").getFormat().autofitColumns();
}