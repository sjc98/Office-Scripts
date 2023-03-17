/**
 * Excel table data can be represented as an array of objects in the form of JSON.
 * Each object represents a row in the table. This helps extract the data from Excel in a consistent format that is visible to the user. The data can then be given to other systems through Power Automate flows.
 */

function main(workbook: ExcelScript.Workbook): TableData[] {
    const table = workbook.getWorksheet("Sheet1").getTables()[0];
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
    "fname": string;
    "Email Address": string;
}
