function main(
  workbook: ExcelScript.Workbook,
  numbersToSum: Array<number> = [],
) {

  let sum = numbersToSum.reduce((a, b) => a + b, 0);
  return sum
}