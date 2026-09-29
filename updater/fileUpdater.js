const AdmZip = require("adm-zip");
const path = require("path");
const fs = require("fs");

/**
 * Replacement dictionary.
 *
 * "Column C": "Column D"
 */
const replacements = {
    "AAGI": "AA",
    "Alpha": "AP",
    "Allstate": "AS",
    "ADS": "AS",
    "APC": "AC",
    "Assurant": "AT",
    "ASSU": "AT",
    "AUL": "AU",
    "Autrella": "AL",
    "AUT": "AL",
    "Carco": "CO",
    "CARS": "CA",
    "Cilajet": "CJ",
    "CILA": "CJ",
    "Crystal Fusion": "CU",
    "CF": "CU",
    "Equity 4 U": "EU",
    "E4U": "EU",
    "GLS": "GL",
    "Gold Standard": "GO",
    "Gold": "GO",
    "GWC": "GC",
    "Headstart": "HW",
    "HWG": "HW",
    "iaAWG": "IW",
    "iaAWG": "IAA",
    "Inside Sales": "IS",
    "Integrity": "IT",
    "INT": "IT",
    "LoJack": "LO",
    "LoJ": "LO",
    "Liberty Shield": "LS",
    "LSC": "LS",
    "Line 5": "LF",
    "L5": "LF",
    "NAAC": "NA",
    "NAC": "NC",
    "NSD": "ND",
    "Other": "OT",
    "Procarma": "PC",
    "PROC": "PC",
    "Riders Advantage": "RA",
    "Riders": "RA",
    "Royal": "RL",
    "ROY": "RL",
    "Shine & Drive": "SD",
    "Shine": "SD",
    "Sonsio": "SO",
    "SON": "SO",
    "USACC": "US",
    "Veritas": "VR",
    "VER": "VR",
    "Wise F&I": "WI",
    "Wise": "WI",
    "Waypoint": "WP",
    "WAY": "WP",
    "Warranty Solutions": "WS",
    "WSS": "WS",
    "Warrantech": "WT",
    "Xzilon": "XI",
    "XZ": "XI",
    "Zero": "ZO",
};

/**
 * Escape XML special characters.
 *
 * This is important because replacement text is inserted back
 * into XML files.
 */
function escapeXml(value) {
    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&apos;");
}

/**
 * Decode the XML entities that can appear inside text nodes.
 *
 * This allows replacements to work against the actual text rather
 * than the encoded XML representation.
 */
function decodeXml(value) {
    return String(value)
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&#39;/g, "'")
        .replace(/&#x27;/gi, "'")
        .replace(/&amp;/g, "&");
}

/**
 * Replace values inside XML <t> text nodes.
 *
 * Excel uses <t> elements for:
 *
 * - Shared strings
 * - Rich text
 * - Inline strings
 *
 * We deliberately ONLY modify the contents of <t>...</t>.
 * This prevents us from accidentally changing XML attributes,
 * formulas, cell references, styles, etc.
 */
/*
function replaceTextNodes(xml, replacementMap) {
    let replacementsMade = 0;

    const updatedXml = xml.replace(
        /(<t(?:\s[^>]*)?>)([\s\S]*?)(<\/t>)/gi,
        (fullMatch, openingTag, textContent, closingTag) => {
            let text = decodeXml(textContent);
            const originalText = text;

            for (const [oldValue, newValue] of Object.entries(replacementMap)) {
                if (text.includes(oldValue)) {
                    text = text.split(oldValue).join(newValue);
                }
            }

            if (text !== originalText) {
                replacementsMade++;
            }

            return `${openingTag}${escapeXml(text)}${closingTag}`;
        }
    );

    return {
        xml: updatedXml,
        replacementsMade,
    };
}
*/

/**
 * Replace values inside XML <t> text nodes.
 *
 * Excel uses <t> elements for:
 *
 * - Shared strings
 * - Rich text
 * - Inline strings
 *
 * We deliberately ONLY modify the contents of <t>...</t>.
 * This prevents us from accidentally changing XML attributes,
 * formulas, cell references, styles, etc.
 * * Fix #1: Case-insensitive regular expression rather than includes()/split().join().
 */
/*
function replaceTextNodes(xml, replacementMap) {
    let replacementsMade = 0;

    const updatedXml = xml.replace(
        /(<t(?:\s[^>]*)?>)([\s\S]*?)(<\/t>)/gi,
        (fullMatch, openingTag, textContent, closingTag) => {
            let text = decodeXml(textContent);

            for (const [oldValue, newValue] of Object.entries(replacementMap)) {
                // Escape regex special characters in the replacement key.
                const escapedOldValue = oldValue.replace(
                    /[.*+?^${}()|[\]\\]/g,
                    "\\$&"
                );

                // Case-insensitive replacement.
                const regex = new RegExp(escapedOldValue, "gi");

                const matches = text.match(regex);

                if (matches) {
                    replacementsMade += matches.length;
                    text = text.replace(regex, newValue);
                }
            }

            return `${openingTag}${escapeXml(text)}${closingTag}`;
        }
    );

    return {
        xml: updatedXml,
        replacementsMade,
    };
}
*/
/**
 * Replace values inside XML <t> text nodes.
 *
 * Excel uses <t> elements for:
 *
 * - Shared strings
 * - Rich text
 * - Inline strings
 *
 * We deliberately ONLY modify the contents of <t>...</t>.
 * This prevents us from accidentally changing XML attributes,
 * formulas, cell references, styles, etc.
 */
function replaceTextNodes(xml, replacementMap) {
    let replacementsMade = 0;

    const updatedXml = xml.replace(
        /(<t(?:\s[^>]*)?>)([\s\S]*?)(<\/t>)/gi,
        (fullMatch, openingTag, textContent, closingTag) => {
            let text = decodeXml(textContent);

            for (const [oldValue, newValue] of Object.entries(replacementMap)) {
                // Escape regex special characters in the replacement key.
                const escapedOldValue = oldValue.replace(
                    /[.*+?^${}()|[\]\\]/g,
                    "\\$&"
                );

                /**
                 * UPDATED REGEX:
                 * ^               = Must be at the start of the string
                 * (?=$| -)        = Must be followed by the end of the string ($) OR " -"
                 */
                const regex = new RegExp(`^${escapedOldValue}(?=$| -)`, "gi");

                const matches = text.match(regex);

                if (matches) {
                    replacementsMade += matches.length;
                    text = text.replace(regex, newValue);
                }
            }

            return `${openingTag}${escapeXml(text)}${closingTag}`;
        }
    );

    return {
        xml: updatedXml,
        replacementsMade,
    };
}

/**
 * Determine whether a ZIP entry is an XML file containing
 * worksheet/cell text that we should inspect.
 */
function shouldProcessXmlEntry(entryName) {
    const normalized = entryName.replace(/\\/g, "/").toLowerCase();

    // Shared strings contain the text used by many Excel cells.
    if (normalized === "xl/sharedstrings.xml") {
        return true;
    }

    // Worksheet XML files contain inline strings and cell data.
    if (
        normalized.startsWith("xl/worksheets/") &&
        normalized.endsWith(".xml")
    ) {
        return true;
    }

    return false;
}

/**
 * Process a single XLSX file.
 *
 * An XLSX file is a ZIP archive containing XML and other files.
 * We modify only the relevant XML entries and then write the
 * archive back out.
 */
async function replaceWorkbookValues(
    inputFile,
    outputFile,
    replacementMap = replacements
) {
    console.log("");
    console.log("========================================");
    console.log("Processing workbook");
    console.log("========================================");
    console.log("Input :", inputFile);
    console.log("Output:", outputFile);

    if (!fs.existsSync(inputFile)) {
        throw new Error(`Input file does not exist: ${inputFile}`);
    }

    // Make sure the output directory exists.
    const outputDir = path.dirname(outputFile);

    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, {
            recursive: true,
        });
    }

    // Read the XLSX as a ZIP archive.
    const zip = new AdmZip(inputFile);

    let totalReplacements = 0;
    let processedXmlFiles = 0;

    const entries = zip.getEntries();

    console.log(`ZIP entries found: ${entries.length}`);

    for (const entry of entries) {
        if (entry.isDirectory) {
            continue;
        }

        const entryName = entry.entryName;

        if (!shouldProcessXmlEntry(entryName)) {
            continue;
        }

        console.log(`Inspecting: ${entryName}`);

        let xml;

        try {
            xml = entry.getData().toString("utf8");
        } catch (error) {
            console.error(`Failed reading ZIP entry: ${entryName}`);
            throw error;
        }

        const result = replaceTextNodes(xml, replacementMap);

        if (result.replacementsMade > 0) {
            console.log(
                `  Replacements in ${entryName}: ${result.replacementsMade}`
            );

            // Replace the ZIP entry with the updated XML.
            zip.updateFile(
                entryName,
                Buffer.from(result.xml, "utf8")
            );

            totalReplacements += result.replacementsMade;
        }

        processedXmlFiles++;
    }

    // Write the modified XLSX.
    zip.writeZip(outputFile);

    console.log("");
    console.log("Workbook saved:", outputFile);
    console.log("XML files inspected:", processedXmlFiles);
    console.log("Total replacement operations:", totalReplacements);
    console.log("");

    return {
        success: true,
        inputFile,
        outputFile,
        replacements: totalReplacements,
    };
}

/**
 * Recursively find XLSX files.
 *
 * This function is available if you later want to enable
 * recursive processing.
 */
function findXlsxFilesRecursively(dir, fileList = []) {
    const files = fs.readdirSync(dir);

    for (const file of files) {
        const filePath = path.join(dir, file);
        const stat = fs.statSync(filePath);

        if (stat.isDirectory()) {
            findXlsxFilesRecursively(filePath, fileList);
        } else if (
            path.extname(file).toLowerCase() === ".xlsx"
        ) {
            fileList.push(filePath);
        }
    }

    return fileList;
}

/**
 * Process every XLSX file in the supplied directory.
 *
 * This preserves the existing Electron IPC interface:
 *
 * startProcess(data.filePath)
 */
async function startProcess(inputPath) {
    const logs = [];

    logs.push("Starting processing...");

    console.log("");
    console.log("========================================");
    console.log("START PROCESS");
    console.log("========================================");
    console.log("Input path:", inputPath);
    console.log("Type:", typeof inputPath);

    /**
     * The Electron folder picker should return:
     *
     * C:\some\folder
     *
     * rather than:
     *
     * ["C:\some\folder"]
     */
    if (typeof inputPath !== "string") {
        throw new TypeError(
            `startProcess expected a string path, but received ${typeof inputPath}`
        );
    }

    if (!fs.existsSync(inputPath)) {
        throw new Error(
            `Input directory does not exist: ${inputPath}`
        );
    }

    const inputStats = fs.statSync(inputPath);

    if (!inputStats.isDirectory()) {
        throw new Error(
            `Input path is not a directory: ${inputPath}`
        );
    }

    /**
     * Find XLSX files directly inside the selected folder.
     *
     * The output folder is automatically excluded because it is
     * a directory, not an XLSX file.
     */
    const xlsxFiles = fs.readdirSync(inputPath).filter((file) => {
        // Ignore Excel temporary files such as ~$Workbook.xlsx
        if (file.startsWith("~$")) {
            return false;
        }

        return path.extname(file).toLowerCase() === ".xlsx";
    });

    console.log(`Found ${xlsxFiles.length} XLSX file(s).`);

    logs.push(`Found ${xlsxFiles.length} XLSX file(s).`);

    if (xlsxFiles.length === 0) {
        logs.push("No XLSX files found.");

        return {
            success: true,
            processedFiles: 0,
            outputDir: path.join(inputPath, "output"),
            logs,
            files: [],
        };
    }

    /**
     * Create:
     *
     * [selected folder]
     *     └── output
     */
    const outputDir = path.join(inputPath, "output");

    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, {
            recursive: true,
        });

        console.log("Created output directory:", outputDir);
    }

    const processedFiles = [];

    /**
     * Main processing loop.
     */
    for (let i = 0; i < xlsxFiles.length; i++) {
        const fileName = xlsxFiles[i];

        const inputFile = path.join(
            inputPath,
            fileName
        );

        const outputFile = path.join(
            outputDir,
            fileName
        );

        console.log("");
        console.log(
            `Processing ${i + 1}/${xlsxFiles.length}: ${fileName}`
        );

        logs.push(
            `Processing ${i + 1}/${xlsxFiles.length}: ${fileName}`
        );

        try {
            const result = await replaceWorkbookValues(
                inputFile,
                outputFile,
                replacements
            );

            processedFiles.push({
                fileName,
                inputFile,
                outputFile,
                replacements: result.replacements,
                success: true,
            });

            logs.push(
                `Completed: ${fileName} (${result.replacements} replacements)`
            );

            console.log(`✓ Completed: ${fileName}`);
        } catch (error) {
            console.error(
                `✗ Failed processing ${fileName}:`,
                error
            );

            logs.push(
                `FAILED: ${fileName}: ${error.message}`
            );

            processedFiles.push({
                fileName,
                inputFile,
                outputFile,
                replacements: 0,
                success: false,
                error: error.message,
            });

            /**
             * Continue processing the remaining files instead
             * of killing the entire batch.
             */
            continue;
        }
    }

    const successfulFiles = processedFiles.filter(
        (file) => file.success
    );

    const failedFiles = processedFiles.filter(
        (file) => !file.success
    );

    logs.push("");
    logs.push("Processing complete.");
    logs.push(`Successful: ${successfulFiles.length}`);
    logs.push(`Failed: ${failedFiles.length}`);

    console.log("");
    console.log("========================================");
    console.log("PROCESSING COMPLETE");
    console.log("========================================");
    console.log(`Successful: ${successfulFiles.length}`);
    console.log(`Failed: ${failedFiles.length}`);
    console.log("Output directory:", outputDir);
    console.log("========================================");
    console.log("");

    return {
        success: failedFiles.length === 0,
        processedFiles: processedFiles.length,
        successfulFiles: successfulFiles.length,
        failedFiles: failedFiles.length,
        outputDir,
        logs,
        files: processedFiles,
    };
}

module.exports = {
    startProcess,
    replaceWorkbookValues,
    findXlsxFilesRecursively,
    replacements,
};
