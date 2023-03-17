function main(workbook: ExcelScript.Workbook) {
	let selectedSheet = workbook.getActiveWorksheet();
	// Insert at range A:D on selectedSheet, move existing cells right
	selectedSheet.getRange("A:D").insert(ExcelScript.InsertShiftDirection.right);
	// Set wrap text to true for range 1:3 on selectedSheet
	selectedSheet.getRange("1:3").getFormat().setWrapText(true);
	// Set wrap text to false for range 1:3 on selectedSheet
	selectedSheet.getRange("1:3").getFormat().setWrapText(false);
	// Set range A3:D3 on selectedSheet
	selectedSheet.getRange("A3:D3").setValues([["Site","Address","Payor Plan","Plan Type"]]);
	// Insert cut cells from J:J on selectedSheet to F:F on selectedSheet.
	selectedSheet.getRange("F:F").insert(ExcelScript.InsertShiftDirection.right);
	selectedSheet.getRange("J:J").moveTo(selectedSheet.getRange("F:F"));
	// Auto fit the columns of range G:J on selectedSheet
	selectedSheet.getRange("G:J").getFormat().autofitColumns();
	// Set width of column(s) F:F on selectedSheet to 144.75
	selectedSheet.getRange("F:F").getFormat().setColumnWidth(144.75);
	// Clear ExcelScript.ClearApplyTo.contents from range F1:F2 on selectedSheet
	selectedSheet.getRange("F1:F2").clear(ExcelScript.ClearApplyTo.contents);
	// Set range H3:J3 on selectedSheet
	selectedSheet.getRange("H3:J3").setValues([["Minimum","Maximum","Median"]]);
	// Set range C4:D4 on selectedSheet
	selectedSheet.getRange("C4:D4").setValue("General");
	// Auto fill range
	selectedSheet.getRange("C4:D4").autoFill();
	// Set range A4 on selectedSheet
	selectedSheet.getRange("A4").setValue("CRCH");
	// Auto fill range
	selectedSheet.getRange("A4").autoFill();
	// Set width of column(s) K:K on selectedSheet to 120.75
	selectedSheet.getRange("K:K").getFormat().setColumnWidth(120.75);
	// Insert at range K:Q on selectedSheet, move existing cells right
	selectedSheet.getRange("K:Q").insert(ExcelScript.InsertShiftDirection.right);
	// Set width of column(s) K:Q on selectedSheet to 69.75
	selectedSheet.getRange("K:Q").getFormat().setColumnWidth(69.75);
	// Set range K3 on selectedSheet
	selectedSheet.getRange("K3").setFormulaLocal("=A$3");
	// Auto fill range
	selectedSheet.getRange("K3").autoFill("K3:Q3", ExcelScript.AutoFillType.fillDefault);
	// Set range K4 on selectedSheet
	selectedSheet.getRange("K4").setFormulaLocal("=$A4");
	// Set range M4:Q4 on selectedSheet
	selectedSheet.getRange("M4:Q4").setFormulasLocal([["=R$1","=R$2","=$E4","=$F4","=$G4"]]);
	// Set number format for O:P on selectedSheet
	selectedSheet.getRange("O:P").setNumberFormatLocal("0");
	// Auto fill range
	selectedSheet.getRange("K4:Q4").autoFill();
	// Insert copied cells from K:Q on selectedSheet to U:AA on selectedSheet.
	selectedSheet.getRange("U:AA").insert(ExcelScript.InsertShiftDirection.right);
	selectedSheet.getRange("U:AA").copyFrom(selectedSheet.getRange("K:Q"));
	// Insert copied cells from K:Q on selectedSheet to AF:AM on selectedSheet.
	selectedSheet.getRange("AF:AM").insert(ExcelScript.InsertShiftDirection.right);
	selectedSheet.getRange("AF:AM").copyFrom(selectedSheet.getRange("K:Q"));
}