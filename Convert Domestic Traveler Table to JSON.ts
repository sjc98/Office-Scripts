/**
 * This is a slim version of the proper way to do this - simply because all values are returned in as a string variable
 * Excel table data can be represented as an array of objects in the form of JSON.
 * Each object represents a row in the table. This helps extract the data from Excel in a consistent format that is visible to the user. The data can then be given to other systems through Power Automate flows.
 */

function main(workbook: ExcelScript.Workbook): TableData[] {
  // Get the first table in the "PlainTable" worksheet.
  // If you know the table name, use `workbook.getTable('TableName')` instead.
  const table = workbook.getWorksheet("Domestic Travelers").getTables()[0];

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
  "Last Name": string;
  "First Name": string;
  "Extension Status": string;
  "Shift": string;
  "36 or 48/wk": string;
  "Carilion Req Number": string;
  "Type of Request": string;
  "CC# (all IP RNs under 0101/6037)": string;
  "Requisition Date": string;
  "Department": string;
  "Position": string;
  "Rate": string;
  "OT BR": string;
  "Staffing Company": string;
  "Start Date": string;
  "End Date": string;
  "Extension/Notes": string;
  "Rate Change": string;
  "ARF Sent": string;
  "CIN": string;
  "DOB": string;
  "Manager (all IP RNs under Ashley Diamond)": string;
  "Email address": string;
  "COVID19 Vaccine Dose 1": string;
  "COVID19 Vaccine Dose 2": string;
  "COVID19 Exemption (Medical/Religious)": string;
  "Domestic or International": string;
  "Unit": string;
  "Facility": string;
  "Level of Care": string;
  "Actual Unit": string;
  "Role-Facility": string;
  "Role-Unit": string;
  "Role and Level of Care": string;
  "Facility, Department Classification, & Role": string;
  "Role and Department Classification": string;
  "Service Line": string;
  "Use Service Line": string;
  "Actual Cost Center": string;
  "DaysOut": string;
  "FullName": string;
  "UNIQID": string;
  "Job Role": string;
  "Vacancy": string;
  "Turnover": string;
  "Overtime": string;
  "Needed": string;
  "Current FTEs": string;
  "Traveler FTEs": string;
  "Functional Vacancy": string;
  "Orientation FTEs": string;
  "ESI Hours": string;
  "Hired Not Started FTEs": string;
  "Vacancy Rate FTEs": string;
  "Open & Posted FTEs": string;
  "On Hold & Posted FTEs": string;
  "Open & Posted Positions": string;
  "On Hold & Posted Positions": string;
  "Approved Openings": string;
  "Current Leave FTEs": string;
  "1ST": string;
  "Unit Director/Manager": string;
  "VP": string;
  "MonthStart": string;
  "MonthEnd": string;
  "First Renewal Date": string;
  "Second Renewal Date": string;
  "Third Renewal Date": string;
  "FY2023": string;
  "Status": string;
  "Completed": string;
  "Department Classification": string;
  "ContractStart": string;
  "ContractEnd": string;
  "Extending AND Offered": string;
  "Shift+": string;
  "Hours": string;
  "FTE": string;
  "Current Rates per day": string;
  "Currently Left on Contract": string;
  "Contract Totals thru Oct 1": string;
  "Contract Worth for 91 days": string;
  "First Renewal Wage": string;
  "Second Renewal Wage": string;
  "Third Renewal Wage": string;
  "PTOStart": string;
  "PTOEnd": string;
  "AllRateDeviation": string;
  "Current&PendingRateDeviation": string;
  "Formatted Rate for Automation": string;
}
