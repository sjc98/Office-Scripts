let table1 = workbook.getActiveWorksheet().getRange("A5").getSurroundingRegion();
let values = table1.getValues();
console.log(values);

let newValues = values.map(row => row.map(cell => cell.toString().trim()));
table1.setValues(newValues);

let table2 = workbook.getActiveWorksheet().getRange("A5").getSurroundingRegion();

console.log(table2.getValues());