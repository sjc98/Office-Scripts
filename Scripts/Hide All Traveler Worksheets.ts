function main(workbook: ExcelScript.Workbook) {
	let detail = workbook.getWorksheet("Detail");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	detail.setVisibility(ExcelScript.SheetVisibility.hidden);
	let vP_Summary = workbook.getWorksheet("VP Summary");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	vP_Summary.setVisibility(ExcelScript.SheetVisibility.hidden);
	let unit_Summary = workbook.getWorksheet("Unit Summary");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	unit_Summary.setVisibility(ExcelScript.SheetVisibility.hidden);
	let hR_Summary = workbook.getWorksheet("HR Summary");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	hR_Summary.setVisibility(ExcelScript.SheetVisibility.hidden);
	let facility_Summary = workbook.getWorksheet("Facility Summary");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	facility_Summary.setVisibility(ExcelScript.SheetVisibility.hidden);
	let covering = workbook.getWorksheet("Covering");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	covering.setVisibility(ExcelScript.SheetVisibility.hidden);
	let key = workbook.getWorksheet("Key");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	key.setVisibility(ExcelScript.SheetVisibility.hidden);
	let vPMetrics = workbook.getTable("VPMetrics");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	vPMetrics.setVisibility(ExcelScript.SheetVisibility.hidden);
	let automations = workbook.getWorksheet("Automations");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	automations.setVisibility(ExcelScript.SheetVisibility.hidden);
	let domestic_Travelers = workbook.getWorksheet("Domestic Travelers");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	domestic_Travelers.setVisibility(ExcelScript.SheetVisibility.hidden);
}