const replacements = {
    "Administrator": "ADMIN",
    "Accounting": "ACCT",
    "Vehicle Protection": "VP",
    "Warranty": "WAR"
};

const ExcelJS = require("exceljs");
const path = require("path");

async function replaceWorkbookValuesTest(inputFile){
    const replacements = {
        "Administrator": "ADMIN",
        "Accounting": "ACCT",
        "Vehicle Protection": "VP",
        "Warranty": "WAR"
    };

    const workbook = new ExcelJS.Workbook();

    await workbook.xlsx.readFile(inputFile);

    // Iterate through every worksheet
    workbook.eachSheet((worksheet) => {
        console.log(`Processing sheet: ${worksheet.name}`);

        // Iterate through every row
        worksheet.eachRow((row) => {

            // Iterate through every cell
            row.eachCell((cell) => {

                /*
                // Only process string cells
                if (typeof cell.value !== "string") {
                    return;
                }

                // Exact match replacement
                if (replacements[cell.value]) {
                    console.log(
                        `Replacing "${cell.value}" -> "${replacements[cell.value]}"`
                    );

                    cell.value = replacements[cell.value];
                }
                */
                // Replace multiple occurrences of the keys in the cell value
                /*
                if (typeof cell.value !== "string") {
                    return;
                }

                let value = cell.value;

                for (const [oldValue, newValue] of Object.entries(replacements)) {
                    value = value.replaceAll(oldValue, newValue);
                }

                cell.value = value;
                */

                // Replace multiple occurrences of the keys in the cell value and preserve text formatting
                // Plain strings
                if (typeof cell.value === "string") {

                    let value = cell.value;

                    for (const [oldValue, newValue] of Object.entries(replacements)) {
                        value = value.replaceAll(oldValue, newValue);
                    }

                    cell.value = value;
                }

                // Rich text
                else if (
                    cell.value &&
                    cell.value.richText
                ) {

                    cell.value.richText.forEach(part => {

                        for (const [oldValue, newValue] of Object.entries(replacements)) {
                            part.text = part.text.replaceAll(oldValue, newValue);
                        }

                    });

                }

            });

        });

    });

    // Save workbook
    await workbook.xlsx.writeFile(inputFile);

    console.log("Workbook saved:", inputFile);

}

module.exports = {
    replaceWorkbookValuesTest
};