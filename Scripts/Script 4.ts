function main(workbook: ExcelScript.Workbook) {
	let selectedSheet = workbook.getActiveWorksheet();
	// Set range F9 on selectedSheet
	selectedSheet.getRange("F9").setFormulaLocal("=CLEAN(TRIM(PROPER(A9)))&\", \"&CLEAN(TRIM(IF(LEN(B9)<6,B9,PROPER(B9))))&\" • \"&CLEAN(TRIM(PROPER(C9)))");
}