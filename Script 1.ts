function main(
  workbook: ExcelScript.Workbook,
  sheetName: string = "Sheet1",
  chartName: string = "Chart 2",
  chartWidth: number = 550): string {
  const chart = workbook.getWorksheet(sheetName).getChart(chartName);
  const chartImage = chart.getImage(chartWidth);
  const result = `data:image/png;base64,${chartImage}`;
  return result;
}
