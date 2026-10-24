const os = require('os');
const crypto = require('crypto');

/**
 * Generates a standard formatted log header string.
 * @param {Object} [config] Optional parameters to override header defaults.
 * @returns {string} Formatted log header string.
 */
function generateLogHeader(config = {}) {
  const {
    appName = 'Application',
    appVersion = '1.0.0',
    environment = process.env.NODE_ENV || 'Development',
    hostName = os.hostname?.() || 'localhost',
    pid = process.pid || 'N/A',
    startTime = new Date().toISOString(),
    runId = crypto.randomUUID?.() || 'N/A',
    logLevel = 'INFO',
    commandArgs = process.argv ? process.argv.slice(2).join(' ') || 'none' : 'none',
    dividerLength = 80
  } = config;

  const divider = '='.repeat(dividerLength);

  return [
    divider,
    'LOG HEADER',
    divider,
    `Application Name : ${appName}`,
    `App Version      : ${appVersion}`,
    `Environment      : ${environment}`,
    `Host Name        : ${hostName}`,
    `Process ID (PID) : ${pid}`,
    `Start Time (UTC) : ${startTime}`,
    `Run ID / Trace   : ${runId}`,
    `Log Level        : ${logLevel}`,
    `Command / Args   : ${commandArgs}`,
    divider,
    '' // Trailing newline separator
  ].join('\n');
}

function generateLogBody(logs) {
  const logLines = logs.map(log => {
    //const timestamp = new Date().toISOString();
    return `[${log.timestamp}] : [${log.level}] ${log.message}`;
  });
  return logLines.join('\n');
}

/**
 * Helper to format milliseconds into HH:MM:SS.mmm format.
 */
function formatDuration(ms) {
  const hours = Math.floor(ms / 3600000);
  const minutes = Math.floor((ms % 3600000) / 60000);
  const seconds = Math.floor((ms % 60000) / 1000);
  const millis = ms % 1000;

  const pad = (n, width = 2) => String(n).padStart(width, '0');
  return `${pad(hours)}:${pad(minutes)}:${pad(seconds)}.${pad(millis, 3)} (${ms}ms)`;
}

/**
 * Generates a standard formatted text log footer string.
 * @param {Object} [config] Footer metrics and configurations.
 * @returns {string} Formatted log footer string.
 */
function generateLogFooter(config = {}) {
  const endTimeObj = config.endTime ? new Date(config.endTime) : new Date();
  const startTimeObj = config.startTime ? new Date(config.startTime) : endTimeObj;
  const durationMs = Math.max(0, endTimeObj - startTimeObj);

  const {
    records = { processed: 0, succeeded: 0, failed: 0 },
    logSummary = { info: 0, warn: 0, error: 0 },
    exitCode = 0,
    exitStatus = exitCode === 0 ? '0 (SUCCESS)' : `${exitCode} (FAILURE)`,
    dividerLength = 80
  } = config;

  const divider = '='.repeat(dividerLength);

  const recordLine = `Processed: ${records.processed ?? 0} | Succeeded: ${records.succeeded ?? 0} | Failed: ${records.failed ?? 0}`;
  const summaryLine = `INFO: ${logSummary.info ?? 0} | WARN: ${logSummary.warn ?? 0} | ERROR: ${logSummary.error ?? 0}`;

  return [
    '', // Leading line break before the footer
    divider,
    'LOG FOOTER',
    divider,
    `End Time (UTC)   : ${endTimeObj.toISOString()}`,
    `Total Execution  : ${formatDuration(durationMs)}`,
    `Total Records    : ${recordLine}`,
    `Log Summary      : ${summaryLine}`,
    `Exit Code / Status: ${exitStatus}`,
    divider
  ].join('\n');
}

module.exports = {
    generateLogHeader,
    generateLogBody,
    generateLogFooter
}