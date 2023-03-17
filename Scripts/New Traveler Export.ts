/**
 * Excel table data can be represented as an array of objects in the form of JSON.
 * Each object represents a row in the table. This helps extract the data from Excel in a consistent format that is visible to the user. The data can then be given to other systems through Power Automate flows.
 */

function main(workbook: ExcelScript.Workbook): TableData[] {
    const table = workbook.getWorksheet("Domestic Travelers").getTables()[0];
    const texts = table.getRange().getTexts();

    // Filter out rows that do not contain any non-empty cells.
    const filteredValues = texts.filter(row => row.some(cell => cell.trim() !== ''));

    let returnObjects: TableData[] = [];
    if (filteredValues.length > 0) {
        returnObjects = returnObjectFromValues(filteredValues);
    }

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

        // Skip the current iteration if the ID field is empty
        if (values[i][0].trim() === '') {
            continue;
        }

        let object = {};
        for (let j = 0; j < values[i].length; j++) {
            // Convert the "Rate", "Hired Not Started FTEs", and "Functional Vacancy" fields to numbers.
            if (objectKeys[j] === "ID" || objectKeys[j] === " Rate " || objectKeys[j] === " OT BR " || objectKeys[j] === "ESI Hours" || objectKeys[j] === "Orientation FTEs" || objectKeys[j] === "Current Leave FTEs" || objectKeys[j] === "Approved Openings" || objectKeys[j] === "FTE" || objectKeys[j] === "Hours" || objectKeys[j] === "Needed" || objectKeys[j] === "Traveler FTEs" || objectKeys[j] === "Current FTEs" || objectKeys[j] === "Hired Not Started FTEs" || objectKeys[j] === "Functional Vacancy") {
                object[objectKeys[j]] = Number(values[i][j]);
            } else {
                object[objectKeys[j]] = values[i][j];
            }
        }

        objectArray.push(object as TableData);
    }

    return objectArray;
}

interface TableData {
    "ID": number;
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
    " Rate ": number;
    " OT BR ": number;
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
    "Needed": number;
    "Current FTEs": number;
    "Traveler FTEs": number;
    "Functional Vacancy": string;
    "Orientation FTEs": number;
    "ESI Hours": number;
    "Hired Not Started FTEs": number;
    "Vacancy Rate FTEs": string;
    "Open & Posted FTEs": number;
    "On Hold & Posted FTEs": string;
    "Open & Posted Positions": string;
    "On Hold & Posted Positions": string;
    "Approved Openings": number;
    "Current Leave FTEs": number;
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
    "Hours": number;
    "FTE": number;
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
