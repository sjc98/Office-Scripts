function main(workbook: ExcelScript.Workbook): ReportImages {
    // Recalculate the workbook to ensure all tables and charts are updated.
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);

    // Get the data from the "InvoiceAmounts" table.
    let sheet1 = workbook.getWorksheet("Sheet1");
    const table = workbook.getWorksheet('InvoiceAmounts').getTables()[0];
    const rows = table.getRange().getTexts();

    // Get only the "Customer Name" and "Amount due" columns, then remove the "Total" row.
    const selectColumns = rows.map((row) => {
        return [row[2], row[5]];
    });
    table.setShowTotals(true);
    selectColumns.splice(selectColumns.length - 1, 1);
    console.log(selectColumns);

    // Delete the "ChartSheet" worksheet if it's present, then recreate it.
    workbook.getWorksheet('ChartSheet')?.delete();
    const chartSheet = workbook.addWorksheet('ChartSheet');

    // Add the selected data to the new worksheet.
    const targetRange = chartSheet.getRange('A1').getResizedRange(selectColumns.length - 1, selectColumns[0].length - 1);
    targetRange.setValues(selectColumns);

    // Insert the chart on sheet 'ChartSheet' at cell "D1".
    let chart_2 = chartSheet.addChart(ExcelScript.ChartType.columnClustered, targetRange);
    chart_2.setPosition('D1');

    // Get images of the chart and table, then return them for a Power Automate flow.
    const chartImage = chart_2.getImage();
    const tableImage = table.getRange().getImage();
    return { chartImage, tableImage };
}

// The interface for table and chart images.
interface ReportImages {
    chartImage: string
    tableImage: string
}