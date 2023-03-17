function main(workbook: ExcelScript.Workbook) {
	let selectedSheet = workbook.getActiveWorksheet();
	// Auto fit the columns of range range A2:H46 on selectedSheet
	selectedSheet.getRange("A2:H46").getFormat().autofitColumns();
}