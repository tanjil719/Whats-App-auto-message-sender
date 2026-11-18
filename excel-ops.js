'use strict';

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');

function loadConfig() {
  const configPath = path.join(__dirname, 'config.json');
  if (!fs.existsSync(configPath)) {
    throw new Error('config.json not found: ' + configPath);
  }
  return JSON.parse(fs.readFileSync(configPath, 'utf8'));
}

function cellText(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((r) => r.text || '').join('');
    if (v.text != null) return String(v.text);
    if (v.result != null) return String(v.result);
    if (v.hyperlink) return String(v.hyperlink);
    return '';
  }
  return String(v);
}

function isExcelFileOpen(excelPath) {
  try {
    const dir = path.dirname(excelPath);
    const base = path.basename(excelPath);
    // Microsoft Office lock file (e.g. "~$file.xlsx").
    if (fs.existsSync(path.join(dir, '~$' + base))) return true;
    // Fallback: try opening for write access.
    const fd = fs.openSync(excelPath, 'r+');
    fs.closeSync(fd);
    return false;
  } catch (err) {
    return true;
  }
}

function getWorksheet(workbook, sheetName) {
  const ws = sheetName ? workbook.getWorksheet(sheetName) : workbook.worksheets[0];
  if (!ws) {
    throw new Error('Worksheet not found. Available sheets: ' +
      workbook.worksheets.map((w) => w.name).join(', '));
  }
  return ws;
}

function buildHeaderMap(ws) {
  const headerRow = ws.getRow(1);
  const map = {};
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const name = cellText(cell).trim();
    if (name) map[name] = colNumber;
  });
  return map;
}

// Removes rows where BOTH the phone number and the message text are identical
// (keeps the first occurrence). This is a duplicate only when the same number
// has the exact same message content.
async function removeDuplicateRows(excelPath) {
  if (!excelPath || !fs.existsSync(excelPath)) {
    throw new Error('Excel file not found: ' + excelPath);
  }
  if (isExcelFileOpen(excelPath)) {
    throw new Error('The Excel file is currently open. Please close it and try again.');
  }

  const config = loadConfig();
  const colPhone = config.columns && config.columns.phone;
  const colMessage = config.columns && config.columns.message;

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);
  const ws = getWorksheet(workbook, config.sheet_name);
  const headerMap = buildHeaderMap(ws);

  const phoneCol = headerMap[colPhone];
  const messageCol = headerMap[colMessage];
  if (!phoneCol || !messageCol) {
    throw new Error('Could not find required columns: ' + colPhone + ', ' + colMessage);
  }

  const seen = new Set();
  const rowsToDelete = [];
  const totalData = Math.max(0, ws.actualRowCount - 1);

  for (let r = 2; r <= ws.actualRowCount; r++) {
    const row = ws.getRow(r);
    const phone = cellText(row.getCell(phoneCol)).trim();
    const message = cellText(row.getCell(messageCol)).trim();
    const key = JSON.stringify([phone, message]);
    if (seen.has(key)) {
      rowsToDelete.push(r);
    } else {
      seen.add(key);
    }
  }

  for (let i = rowsToDelete.length - 1; i >= 0; i--) {
    ws.spliceRows(rowsToDelete[i], 1);
  }

  await workbook.xlsx.writeFile(excelPath);
  return {
    total: totalData,
    removed: rowsToDelete.length,
    remaining: totalData - rowsToDelete.length,
  };
}

// Sorts data rows A-Z by the given column (header name), preserving the header.
async function sortByColumn(excelPath, columnName) {
  if (!excelPath || !fs.existsSync(excelPath)) {
    throw new Error('Excel file not found: ' + excelPath);
  }
  if (isExcelFileOpen(excelPath)) {
    throw new Error('The Excel file is currently open. Please close it and try again.');
  }

  const config = loadConfig();

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);
  const ws = getWorksheet(workbook, config.sheet_name);
  const headerMap = buildHeaderMap(ws);

  const sortCol = headerMap[columnName];
  if (!sortCol) {
    throw new Error('Column not found: ' + columnName);
  }

  const totalData = Math.max(0, ws.actualRowCount - 1);
  if (totalData <= 1) {
    return { total: totalData, sorted: totalData };
  }

  const headerValues = ws.getRow(1).values.slice();

  const dataRows = [];
  for (let r = 2; r <= ws.actualRowCount; r++) {
    const row = ws.getRow(r);
    const key = cellText(row.getCell(sortCol)).trim().toLowerCase();
    dataRows.push({ key, values: row.values.slice() });
  }

  dataRows.sort((a, b) =>
    a.key.localeCompare(b.key, undefined, { numeric: true, sensitivity: 'base' })
  );

  ws.spliceRows(1, ws.actualRowCount);

  const headerRow = ws.getRow(1);
  for (let c = 1; c < headerValues.length; c++) {
    if (headerValues[c] !== undefined) headerRow.getCell(c).value = headerValues[c];
  }

  dataRows.forEach((d, i) => {
    const row = ws.getRow(i + 2);
    for (let c = 1; c < d.values.length; c++) {
      if (d.values[c] !== undefined) row.getCell(c).value = d.values[c];
    }
  });

  await workbook.xlsx.writeFile(excelPath);
  return { total: totalData, sorted: totalData };
}

module.exports = {
  isExcelFileOpen,
  removeDuplicateRows,
  sortByColumn,
};
