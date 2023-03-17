function main(workbook: ExcelScript.Workbook) {
	// Unknown event received with eventId:577
	let selectedSheet = workbook.getActiveWorksheet();
	// Create a table with format on range A8:G16339 on selectedSheet
	let table1 = workbook.addTable(selectedSheet.getRange("A8:G16339"), true);
	table1.setPredefinedTableStyle("TableStyleLight1");
	// Unknown event received with eventId:1009
	// Clear fill color for range A8:G16339 on selectedSheet
	selectedSheet.getRange("A8:G16339").getFormat().getFill().clear();
	// Set range F8:G9 on selectedSheet
	selectedSheet.getRange("F8").setFormulasLocal([["Q1"], ["=\"'\" & CLEAN(TRIM(PROPER(A9))) &\", \"& CLEAN(TRIM(IF(LEN(B9)<6,B9,PROPER(B9))))&\" • \"&CLEAN(TRIM(PROPER(C9))) &\"',\"",null]]);
	// Auto fill range
	selectedSheet.getRange("F9").autoFill();
	// Replace all ' with  on range A:A on selectedSheet
	selectedSheet.getRange("A:A").replaceAll("'", "", { completeMatch: false, matchCase: false });
}