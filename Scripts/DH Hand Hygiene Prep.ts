function main(workbook: ExcelScript.Workbook) {
	let table1 = workbook.getTable("Table1");
	// Sort on table: table1 column index: '3'
	table1.getSort().apply([{key: 3, ascending: false}]);
	// Sort on table: table1 column index: '0'
	table1.getSort().apply([{key: 0, ascending: true}]);
	// Toggle auto filter on table table1
	table1.getAutoFilter().apply("D6");
	// Sort on table: table1 column index: '3'
	table1.getSort().apply([{key: 3, ascending: true}]);
	// Toggle auto filter on table table1
	table1.getAutoFilter().remove();
	let oP_Observed = workbook.getWorksheet("OP Observed");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	oP_Observed.setVisibility(ExcelScript.SheetVisibility.hidden);
	let cNRV_IP_Observed = workbook.getWorksheet("CNRV IP Observed");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	cNRV_IP_Observed.setVisibility(ExcelScript.SheetVisibility.hidden);
	let cRMH_IP_Observed = workbook.getWorksheet("CRMH IP Observed");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	cRMH_IP_Observed.setVisibility(ExcelScript.SheetVisibility.hidden);
	let observing_Department = workbook.getWorksheet("Observing Department");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	observing_Department.setVisibility(ExcelScript.SheetVisibility.hidden);
	let compliance = workbook.getWorksheet("Compliance");
	// Set sheet visibility to ExcelScript.SheetVisibility.hidden
	compliance.setVisibility(ExcelScript.SheetVisibility.hidden);
}