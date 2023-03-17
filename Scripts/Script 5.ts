function main(workbook: ExcelScript.Workbook) {
	let summary = workbook.getWorksheet("Summary");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	summary.setVisibility(ExcelScript.SheetVisibility.hidden);
}