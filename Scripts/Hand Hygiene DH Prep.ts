function main(workbook: ExcelScript.Workbook) {
	let data = workbook.getWorksheet("Data");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	data.setVisibility(ExcelScript.SheetVisibility.hidden);
	let dept__Observers = workbook.getWorksheet("Dept. Observers");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	dept__Observers.setVisibility(ExcelScript.SheetVisibility.hidden);
	let summary = workbook.getWorksheet("Summary");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	summary.setVisibility(ExcelScript.SheetVisibility.hidden);
	let selectedSheet = workbook.getActiveWorksheet();
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	selectedSheet.setVisibility(ExcelScript.SheetVisibility.hidden);
	let cNRV_IP_Observed = workbook.getWorksheet("CNRV IP Observed");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	cNRV_IP_Observed.setVisibility(ExcelScript.SheetVisibility.hidden);
	let cRMH_IP_Observed = workbook.getWorksheet("CRMH IP Observed");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	cRMH_IP_Observed.setVisibility(ExcelScript.SheetVisibility.hidden);
	let compliance = workbook.getWorksheet("Compliance");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	compliance.setVisibility(ExcelScript.SheetVisibility.hidden);
	let table1 = workbook.getTable("Table1");
	// Reapply filter on table table1
	table1.reapplyFilters();
}