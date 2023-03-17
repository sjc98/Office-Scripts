function main(workbook: ExcelScript.Workbook) {
	let patient_Days = workbook.getWorksheet("Patient Days");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	falls_Data.setVisibility(ExcelScript.SheetVisibility.hidden);
	let falls_Data = workbook.getWorksheet("Falls Data");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	year_Summary.setVisibility(ExcelScript.SheetVisibility.hidden);
	let year_Summary = workbook.getWorksheet("Year Summary");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	year_Summary.setVisibility(ExcelScript.SheetVisibility.hidden);
}