function main(
    workbook: ExcelScript.Workbook,
    sheetName: string = "Sheet1",
    chartName: string = "Chart 8",
    chartWidth: number = 220): string {
    const chart = workbook.getWorksheet(sheetName).getChart(chartName);
    const chartImage = chart.getImage(chartWidth);
    const result = `data:image/png;base64,${chartImage}`;
    return result;
}