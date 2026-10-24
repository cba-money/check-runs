const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');

const { generateLogHeader, generateLogBody, generateLogFooter } = require('./logs.js');

function fileExists(filePath) {
  return fs.existsSync(filePath) && fs.statSync(filePath).isFile();
}

function findColumnByHeader(worksheet, possibleNames) {
  const headerRow = worksheet.getRow(1);
  const normalized = possibleNames.map(h => h.toLowerCase().replace(/[^a-z0-9]/g, ''));

  for (let i = 1; i <= headerRow.cellCount; i++) {
    const cellValue = String(headerRow.getCell(i).value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    if (normalized.includes(cellValue)) return i;
  }
  return -1;
}

function detectHeaderRow(ws) {
  const row1 = ws.getRow(1).values.join(' ').toLowerCase();
  const row2 = ws.getRow(2).values.join(' ').toLowerCase();
  const hasHeaders = str =>
    str.includes('check') && (str.includes('debit') || str.includes('amount'));
  return hasHeaders(row1) ? 1 : hasHeaders(row2) ? 2 : 1;
}

function normalizeValue(value) {
  if (value == null) return '';
  if (typeof value === 'object' && value.text) return String(value.text).trim();
  if (typeof value === 'object' && value.richText) return value.richText.map(x => x.text).join('').trim();
  return String(value).trim();
}

function getDisplayDate(value) {
  if (!value) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'object' && value.result instanceof Date) return value.result.toISOString().slice(0, 10);
  return String(value).trim();
}

function generateLogs(logs, startTime, endTime, exitCode, exitStatus) {
  const header = generateLogHeader({
    appName: 'Check Runs Desktop',
    appVersion: '2.4.1',
    environment: 'Production',
    logLevel: 'DEBUG',
    commandArgs: 'node server.js --port=8080'
  });

  const logBody = generateLogBody(logs);

  const recordsProcessed = logs.length;
  const recordsSucceeded = logs.filter(log => log.level === 'SUCCESS').length;
  const recordsFailed = logs.filter(log => log.level === 'ERROR').length;

  const footer = generateLogFooter({
    startTime: startTime,
    endTime: endTime,
    records: {
      processed: recordsProcessed,
      succeeded: recordsSucceeded,
      failed: recordsFailed
    },
    logSummary: {
      info: logs.filter(log => log.level === 'INFO').length,
      warn: logs.filter(log => log.level === 'WARN').length,
      error: logs.filter(log => log.level === 'ERROR').length
    },
    exitCode: 0,
    exitStatus: '0 (SUCCESS_WITH_WARNINGS)'
  });

  return [header, logBody, footer].join('\n');
}

async function startProcessing(checkRegisterPath, runPath, outputPath, saveLogs) {
  const startTime = new Date(Date.now());
  let logs = [{
    message: `Starting processing with Register: ${checkRegisterPath}, Run: ${runPath}, Output Path: ${outputPath}, Save Logs: ${saveLogs}`,
    level: 'INFO',
    timestamp: new Date().toISOString()
  }];
  const registerFileExists = fileExists(checkRegisterPath);
  const runFileExists = fileExists(runPath);
  const outputDirExists = fs.existsSync(outputPath) && fs.statSync(outputPath).isDirectory();

  if (!registerFileExists || !runFileExists) {
    logs.push({
      message: 'One or more specified files do not exist.',
      level: 'ERROR',
      timestamp: new Date().toISOString()
    });
    throw new Error('One or more specified files do not exist.');
  }

  if(!outputDirExists) {
    logs.push({
      message: 'The specified output path is not a valid directory.',
      level: 'ERROR',
      timestamp: new Date().toISOString()
    });
    throw new Error('The specified output path is not a valid directory.');
  }

  const registerWb = new ExcelJS.Workbook();
  const runWb = new ExcelJS.Workbook();

  await registerWb.xlsx.readFile(checkRegisterPath);
  await runWb.xlsx.readFile(runPath);
  logs.push({
    message: 'Successfully read both Excel files.',
    level: 'INFO',
    timestamp: new Date().toISOString()
  });

  const registerWs = registerWb.worksheets[0];
  const runWs = runWb.worksheets[0];

  //Remove top row if its not the header
  const runHeaderRow = detectHeaderRow(runWs);
  if (runHeaderRow === 2) runWs.spliceRows(1, 1);

  const registerCheckCol = findColumnByHeader(registerWs, ['Check #', 'Check Number']);
  let registerDateCashedCol = findColumnByHeader(registerWs, ['Date Cashed?', 'Date Cashed']);
  const registerAmountCol = findColumnByHeader(registerWs, ['Amount', 'Check Amount']);

  if (registerDateCashedCol === -1) {
    registerDateCashedCol = registerWs.columnCount + 1;
    registerWs.getRow(1).getCell(registerDateCashedCol).value = 'Date Cashed?';
  }

  const runCheckCol = findColumnByHeader(runWs, ['Check #', 'Check Number', 'Cust. Ref. / Check Number', 'Check']);
  const runAmountCol = findColumnByHeader(runWs, ['Debit', 'Amount']);
  const runDateCol = findColumnByHeader(runWs, ['Date', 'Check Date', 'Date Cashed', 'Post Date']);

  if ([registerCheckCol, registerDateCashedCol, registerAmountCol, runCheckCol, runAmountCol].includes(-1)) {
    logs.push({
      message: 'One or more required headers were not found in the Excel files.',
      level: 'ERROR',
      timestamp: new Date().toISOString()
    });
    throw new Error('One or more required headers were not found.');
  }

  logs.push({
    message: 'All required headers found.',
    level: 'INFO',
    timestamp: new Date().toISOString()
  });

  const discrepancies = [];

  const normalFont = {
    bold: false,
    color: { argb: 'FF000000' }
  };

  const yellowFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF00' }
  };

  const blueFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFADD8E6' }
  };

  const voidFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FFFFFF00' }
  };

  const voidFont = {
    bold: true,
    color: { argb: 'FFFF0000' }
  };

  const voidedChecks = [];
  const voidedBases = new Set();

  function styleRow(row, fill) {
    if (row.number === 1) return;

    row.eachCell(cell => {
      cell.fill = fill;
    });
  }

  function getRegisterRowByCheck(checkNumber) {
    for (let rr = 2; rr <= registerWs.rowCount; rr++) {
      const regRow = registerWs.getRow(rr);
      const regCheck = normalizeValue(regRow.getCell(registerCheckCol).value);
      if (regCheck === checkNumber) return regRow;
    }
    return null;
  }

  for (let rr = 2; rr <= registerWs.rowCount; rr++) {
    const regRow = registerWs.getRow(rr);
    const regCheck = normalizeValue(regRow.getCell(registerCheckCol).value);

    if (!regCheck || !regCheck.endsWith('V')) continue;

    if (regCheck.endsWith('V')) {
      const voidRow = getRegisterRowByCheck(regCheck);
      const voidCell = voidRow.getCell(registerDateCashedCol);
      voidCell.value = "VOID";
      voidCell.alignment = { horizontal: 'right' };
      voidCell.font = {
        bold: true,
        color: { argb: 'FFFF0000' }
      };

      const baseCheck = regCheck.slice(0, -1);
      voidedChecks.push(regCheck);
      voidedBases.add(baseCheck);

      const duplicateRow = getRegisterRowByCheck(baseCheck);
      if (!duplicateRow) continue;

      const dateCell = duplicateRow.getCell(registerDateCashedCol);
      dateCell.value = 'VOID';
      dateCell.alignment = { horizontal: 'right' };
      dateCell.font = {
        bold: true,
        color: { argb: 'FFFF0000' }
      };

      logs.push({
        message: `Voided check detected: ${regCheck}. Marked both ${regCheck} and ${baseCheck} as VOID.`,
        level: 'INFO',
        timestamp: new Date().toISOString()
      });
    }
  }

  for (let r = 2; r <= runWs.rowCount; r++) {
    const runRow = runWs.getRow(r);
    let hasDiscrepancy = false;
    const checkNumber = normalizeValue(runRow.getCell(runCheckCol).value);
    const runAmount = Number(normalizeValue(runRow.getCell(runAmountCol).value).replace(/[^0-9.-]/g, ''));

    if (!checkNumber) continue;

    if (voidedBases.has(checkNumber)) {
      styleRow(runRow, blueFill);
      hasDiscrepancy = true;
      discrepancies.push({
        checkNumber,
        issue: 'Voided check appears in check run',
        row: r
      });
      continue;
    }

    const matches = [];
    
    for (let rr = 2; rr <= registerWs.rowCount; rr++) {
      const regRow = registerWs.getRow(rr);
      const regCheck = normalizeValue(regRow.getCell(registerCheckCol).value);
      if (regCheck === checkNumber) matches.push(regRow);
    }

    // Unknown/Unregistered check cashed discrepancy
    if (matches.length === 0) {
      styleRow(runRow, blueFill);
      hasDiscrepancy = true;
      discrepancies.push({
        checkNumber,
        issue: 'Not found in register',
        row: r
      });
      continue;
    }

    const regRow = matches[0];
    const dateCashed = getDisplayDate(regRow.getCell(registerDateCashedCol).value);
    const regAmount = Number(normalizeValue(regRow.getCell(registerAmountCol).value).replace(/[^0-9.-]/g, ''));

    // Voided check cashed discrepancy
    if (dateCashed === 'VOID') {
      styleRow(runRow, blueFill);
      hasDiscrepancy = true;
      discrepancies.push({
        checkNumber,
        issue: 'Voided check appears in check run',
        row: r
      });
      continue;
    }

    if (!dateCashed) {
      const dateCell = regRow.getCell(registerDateCashedCol);
      const runDate = runRow.getCell(runDateCol).value;

      // Record the cashing date regardless of amount mismatch
      dateCell.value = runDate;
      dateCell.alignment = { horizontal: 'right' };
      dateCell.font = normalFont;

      // Only consider it a normal match if the amounts also match
      if (runAmount === regAmount) {
        // No discrepancy
      }
    }

    // Double/Triple/N (N>1) times cashed discrepancy
    if (dateCashed && dateCashed !== 'VOID') {
      hasDiscrepancy = true;
      styleRow(runRow, blueFill);
      discrepancies.push({
        checkNumber,
        issue: `Previously cashed: ${dateCashed}`,
        row: r
      });
    }

    // Amount mismatch discrepancy
    if (!Number.isNaN(runAmount) && !Number.isNaN(regAmount) && runAmount !== regAmount) {
      hasDiscrepancy = true;
      styleRow(runRow, blueFill);
      discrepancies.push({
        checkNumber,
        issue: `Amount mismatch. Run: ${runAmount}, Register: ${regAmount}`,
        row: r
      });
    }

    if (!hasDiscrepancy) {
      styleRow(runRow, yellowFill);
    }

  }

  const discWb = new ExcelJS.Workbook();
  const discWs = discWb.addWorksheet('Discrepancies');

  discWs.columns = [
    { header: 'Row', key: 'row', width: 10 },
    { header: 'Check Number', key: 'checkNumber', width: 20 },
    { header: 'Issue', key: 'issue', width: 60 }
  ];

  discWs.addRows(discrepancies);

  //const exportDir = path.join(process.cwd(), 'exports');
  const exportDir = path.join(outputPath);

  if (!fs.existsSync(exportDir)) {
    fs.mkdirSync(exportDir, { recursive: true });
  }

  const timestamp = Date.now();
  const endTime = new Date(Date.now());

  const modifiedRunPath = path.join(
    exportDir,
    `modified-check-run-${timestamp}.xlsx`
  );

  const discrepanciesPath = path.join(
    exportDir,
    `discrepancies-${timestamp}.xlsx`
  );

  const modifiedRegisterPath = path.join(
    exportDir,
    `modified-check-register-${timestamp}.xlsx`
  );

  await runWb.xlsx.writeFile(modifiedRunPath);
  await discWb.xlsx.writeFile(discrepanciesPath);
  await registerWb.xlsx.writeFile(modifiedRegisterPath);

  logs.push({
    message: 'Successfully saved all files.',
    level: 'INFO',
    timestamp: new Date().toISOString()
  });


  if(saveLogs) {
    const logFilePath = path.join(exportDir, `processing-log-${timestamp}.txt`);
    const logContent = generateLogs(logs, startTime, endTime, 0, '0 (SUCCESS)');
    fs.writeFileSync(logFilePath, logContent);
  }
}

module.exports = startProcessing;