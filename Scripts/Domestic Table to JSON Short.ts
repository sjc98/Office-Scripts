/**
 * Excel table data can be represented as an array of objects in the form of JSON.
 * Each object represents a row in the table. This helps extract the data from Excel in a consistent format that is visible to the user. The data can then be given to other systems through Power Automate flows.
 */

function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "PlainTable" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getTable('SarahsTable');

  // Get all the values from the table as text.
  const texts = table.getRange().getTexts();

  // Create an array of JSON objects that match the row structure.
  let returnObjects: TableData[] = [];
  if (table.getRowCount() > 0) {
    returnObjects = returnObjectFromValues(texts);
  }

  // Log the information and return it for a Power Automate flow.
  console.log(JSON.stringify(returnObjects));
  return returnObjects;
}

// This function converts a 2D array of values into a generic JSON object.
// In this case, we have defined the TableData object, but any similar interface would work.
function returnObjectFromValues(values: string[][]): TableData[] {
  let objectArray: TableData[] = [];
  let objectKeys: string[] = [];
  for (let i = 0; i < values.length; i++) {
    if (i === 0) {
      objectKeys = values[i];
      continue;
    }

    let object = {};
    for (let j = 0; j < values[i].length; j++) {
      object[objectKeys[j]] = values[i][j];
    }

    objectArray.push(object as TableData);
  }

  return objectArray;
}

interface TableData {
  "Extension Status": string;
  "Shift": string;
  "Carilion Req Number": string;
  "CC# (all IP RNs under 0101/6037)": string;
  "Requisition Date": string;
  "Department": string;
  "Position": string;
  "Rate": string;
  "OT BR": string;
  "Staffing Company": string;
  "Start Date": string;
  "End Date": string;
  "CIN": string;
  "Manager (all IP RNs under Ashley Diamond)": string;
  "Unit": string;
  "Facility": string;
  "Level of Care": string;
  "Actual Unit": string;
  "Role-Facility": string;
  "Service Line": string;
  "Actual Cost Center": string;
  "UNIQID": string;
  "Job Role": string;
  "Unit Director/Manager": string;
  "VP": string;
  "Status": string;
  "Hours": string;
  "FTE": string;
  "ID": string;
}
