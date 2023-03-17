function main(workbook: ExcelScript.Workbook) {
	let selectedSheet = workbook.getActiveWorksheet();
	// Replace all ' with  on range A:A on selectedSheet
	selectedSheet.getRange("A:A").replaceAll("'", "", {completeMatch: false, matchCase: false});
}