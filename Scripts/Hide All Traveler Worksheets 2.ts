function main(workbook: ExcelScript.Workbook) {
	let facility_Summary = workbook.getWorksheet("Facility Summary");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	facility_Summary.setVisibility(ExcelScript.SheetVisibility.hidden);
	let hR_Summary = workbook.getWorksheet("HR Summary");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	hR_Summary.setVisibility(ExcelScript.SheetVisibility.hidden);
	let unit_Summary = workbook.getWorksheet("Unit Summary");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	unit_Summary.setVisibility(ExcelScript.SheetVisibility.hidden);
	let vP_Summary = workbook.getWorksheet("VP Summary");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	vP_Summary.setVisibility(ExcelScript.SheetVisibility.hidden);
	let detail = workbook.getWorksheet("Detail");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	detail.setVisibility(ExcelScript.SheetVisibility.hidden);
}