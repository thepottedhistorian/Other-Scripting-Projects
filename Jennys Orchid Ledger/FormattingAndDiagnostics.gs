/**
 * =============================================================================
 * JENNY'S ORCHID LEDGER - FORMATTING, DATES & DIAGNOSTICS
 * =============================================================================
 * Purpose: Consolidates bloom status styling, duration calculations, repot
 *          stoplight coloring, date standardization, Excel export, and health checks.
 * =============================================================================
 */

/* --- 1. BLOOM STATUS STYLING & DURATION ENGINE --- */

function applyBloomStatusFormatting(sheet) {
  const cell = sheet.getRange("D5");
  let val = cell.getValue().toString().trim();
  let cleanStatus = val.replace(/^[^a-zA-Z0-9]+/, "").trim();
  const statusLower = cleanStatus.toLowerCase();

  let icon = "⬜", bgColor = "#fca5a5", fontColor = "#ffffff";

  if (statusLower.includes("not in bloom")) {
    icon = "⬜"; bgColor = "#fca5a5"; fontColor = "#ffffff";
  } else if (statusLower.includes("in spike") || statusLower.includes("spiking")) {
    icon = "🌱"; bgColor = "#bbf7d0"; fontColor = "#065f46";
  } else if (statusLower.includes("in bud")) {
    icon = "🌼"; bgColor = "#e9d5ff"; fontColor = "#581c87";
  } else if (statusLower.includes("in bloom")) {
    icon = "🌸"; bgColor = "#10b981"; fontColor = "#ffffff";
  }

  cell.setBackground(bgColor).setFontColor(fontColor).setFontWeight("bold").setHorizontalAlignment("center");
  const newString = `${icon} ${cleanStatus}`;
  if (val !== newString) cell.setValue(newString);
}

function refreshAllBloomStatusFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const idPattern = /^\d+(-\d+)?$/;
  ss.getSheets().forEach(sheet => {
    if (idPattern.test(sheet.getName().trim())) applyBloomStatusFormatting(sheet);
  });
}

function updateSheetBloomDurations(sheet) {
  if (!sheet || sheet.getLastRow() < 20) return 0;
  const startRow = 20;
  const numRows = Math.min(sheet.getLastRow() - 19, 30);
  const range = sheet.getRange(startRow, 1, numRows, 2);
  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  let updatedCount = 0;

  for (let i = 0; i < values.length; i++) {
    const startStr = String(displayValues[i][0]).replace(/\u00A0/g, ' ').trim();
    const endStr = String(displayValues[i][1]).replace(/\u00A0/g, ' ').trim();
    const durationText = calculateDurationString_(values[i][0], values[i][1], startStr, endStr);

    const cCell = sheet.getRange(startRow + i, 3);
    if (cCell.getDisplayValue().trim() !== durationText) {
      cCell.setValue(durationText).setHorizontalAlignment("center");
      updatedCount++;
    }
  }
  return updatedCount;
}

function updateAllBloomDurations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const idPattern = /^\d+(-\d+)?$/;
  let totalRowsUpdated = 0;
  let processedSheets = 0;

  ss.getSheets().forEach(sheet => {
    if (idPattern.test(sheet.getName().trim())) {
      totalRowsUpdated += updateSheetBloomDurations(sheet);
      processedSheets++;
    }
  });

  try {
    SpreadsheetApp.getUi().alert(`Bloom Duration Sync Complete!\nAudited ${processedSheets} ID sheets. Updated ${totalRowsUpdated} records.`);
  } catch (e) {}
}

function calculateDurationString_(startDateVal, endDateVal, startStr = "", endStr = "") {
  const invalidTokens = ['', '-', '—', 'null', 'unknown'];
  const sStr = (startStr || String(startDateVal || '')).toLowerCase().trim();
  const eStr = (endStr || String(endDateVal || '')).toLowerCase().trim();
  if (invalidTokens.includes(sStr) || invalidTokens.includes(eStr)) return "";

  let d1 = parseToUtcDate_(startDateVal, startStr);
  let d2 = parseToUtcDate_(endDateVal, endStr);
  if (!d1 || !d2 || d2 < d1) return "";

  let years = d2.getUTCFullYear() - d1.getUTCFullYear();
  let months = d2.getUTCMonth() - d1.getUTCMonth();
  let days = d2.getUTCDate() - d1.getUTCDate();

  if (days < 0) {
    months -= 1;
    days += new Date(Date.UTC(d2.getUTCFullYear(), d2.getUTCMonth(), 0)).getUTCDate();
  }
  if (months < 0) {
    years -= 1;
    months += 12;
  }

  const totalMonths = (years * 12) + months;
  const parts = [];
  if (totalMonths > 0) parts.push(totalMonths + (totalMonths === 1 ? " month" : " months"));
  if (days > 0) parts.push(days + (days === 1 ? " day" : " days"));
  return parts.length > 0 ? parts.join(" ") : "0 days";
}

function parseToUtcDate_(dateVal, dateStr) {
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    return new Date(Date.UTC(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate()));
  }
  const parsed = Date.parse(dateStr);
  if (!isNaN(parsed)) {
    const d = new Date(parsed);
    return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  }
  return null;
}

/* --- 2. REPOTTING STOPLIGHTS --- */

function applyAllSystemStoplights() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  if (!inv) return;

  const data = inv.getDataRange().getValues();
  const today = new Date();
  let count = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const id = data[i][ID_COL];
    if (inv.isRowHiddenByUser(row) || !id) continue;

    const nextDue = data[i][NEXT_COL];
    let color = "#10b981"; // Green

    if (nextDue instanceof Date && !isNaN(nextDue.getTime())) {
      const diffDays = (nextDue - today) / (1000 * 60 * 60 * 24);
      if (diffDays < 0) color = "#ef4444";       // Red (Overdue)
      else if (diffDays < 30) color = "#f59e0b"; // Amber (< 30 days)
    }

    inv.getRange(row, ID_COL + 1).setBackground(color).setFontColor("white");

    const orchidSheet = ss.getSheetByName(String(id).trim());
    if (orchidSheet) {
      orchidSheet.getRange("D10").setBackground(color).setFontColor("white");
      count++;
    }
  }

  try {
    SpreadsheetApp.getUi().alert(`Stoplight Sync Complete: ${count} ledgers updated.`);
  } catch (e) {}
}

/* --- 3. DATE STANDARDIZATION & TIMESTAMP CLEANUP --- */

function standardizeSelectedDates() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Please select a range of cells first.');
    return;
  }
  const count = processDateRange(range);
  SpreadsheetApp.getUi().alert(`Date Standardization Complete!\nUpdated ${count} date entries in selection.`);
}

function standardizeActiveSheetDates() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const count = processDateRange(sheet.getDataRange());
  SpreadsheetApp.getUi().alert(`Date Standardization Complete!\nUpdated ${count} date entries on sheet: "${sheet.getName()}".`);
}

function standardizeAllOrchidInventoryDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Batch Date Standardization',
    'This will standardize dates across Inventory, Archived Orchids, and all Orchid ID sheets.\n\nProceed?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  let totalUpdated = 0;
  let processedSheetsCount = 0;
  const idPattern = /^\d+(-\d+)?$/;

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName().trim();
    if (name === INVENTORY_SHEET || name === ARCHIVE_SHEET || idPattern.test(name)) {
      totalUpdated += processDateRange(sheet.getDataRange());
      processedSheetsCount++;
    }
  });

  ui.alert(`System Batch Complete\nStandardized ${totalUpdated} dates across ${processedSheetsCount} sheets.`);
}

function processDateRange(range) {
  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const sheet = range.getSheet();
  const startRow = range.getRow();
  const startCol = range.getColumn();
  let formattedCount = 0;

  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[0].length; c++) {
      const cellValue = values[r][c];
      const displayText = String(displayValues[r][c]).replace(/\u00A0/g, ' ').trim();
      if (displayText === '' || /^\d{4}$/.test(displayText)) continue;

      let formattedDateString = null;
      const shortDateMatch = displayText.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
      if (shortDateMatch) {
        const monthIdx = parseInt(shortDateMatch[1], 10) - 1;
        const day = parseInt(shortDateMatch[2], 10);
        const year = shortDateMatch[3];
        if (monthIdx >= 0 && monthIdx < 12) {
          formattedDateString = `${monthNames[monthIdx]} ${day}, ${year}`;
        }
      }

      if (!formattedDateString) {
        let dateObj = (cellValue instanceof Date && !isNaN(cellValue.getTime())) ? cellValue : null;
        if (!dateObj) {
          const parsed = Date.parse(displayText);
          if (!isNaN(parsed)) dateObj = new Date(parsed);
        }
        if (dateObj) {
          formattedDateString = Utilities.formatDate(dateObj, timeZone, "MMMM d, yyyy");
        }
      }

      if (formattedDateString && displayText !== formattedDateString) {
        sheet.getRange(startRow + r, startCol + c).setValue("'" + formattedDateString);
        formattedCount++;
      }
    }
  }
  return formattedCount;
}

function fixAllPhasesSequentially() {
  const sheet = SpreadsheetApp.getActiveSheet();
  fixDateRange_(sheet.getRange(20, 1, 20, 2), "mmmm d, yyyy");
  fixDateRange_(sheet.getRange(20, 6, 20, 1), "mmmm d, yyyy");
  fixDateRange_(sheet.getRange(20, 9, 20, 1), "mmmm d, yyyy");
  fixDateRange_(sheet.getRange(20, 12, 40, 1), "mmmm d, yyyy");
  SpreadsheetApp.getUi().alert("Timestamp formatting fixed for: " + sheet.getName());
}

function fixDateRange_(range, format) {
  const values = range.getValues();
  let changed = false;
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[0].length; c++) {
      const v = values[r][c];
      if (v instanceof Date) continue;
      if (typeof v === "string") {
        const cleaned = v.replace(/[^\S\r\n]+/g, " ").replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
        if (!cleaned) continue;
        const parsed = new Date(cleaned);
        if (!isNaN(parsed.getTime())) {
          values[r][c] = parsed;
          changed = true;
        }
      }
    }
  }
  if (changed) range.setValues(values);
  range.setHorizontalAlignment("left").setNumberFormat(format);
}

/* --- 4. EXCEL EXPORT & SYSTEM HEALTH CHECK --- */

function exportWorkbookToXLSX() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const originalName = ss.getName();
  const tempFile = DriveApp.getFileById(ss.getId()).makeCopy(`${originalName}_Excel_Export`);
  const tempSs = SpreadsheetApp.openById(tempFile.getId());

  tempSs.getSheets().forEach(sheet => {
    const maxRows = sheet.getMaxRows();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    sheet.clearConditionalFormatRules();
    if (lastRow > 0 && lastCol > 0) {
      const range = sheet.getRange(1, 1, lastRow, lastCol);
      range.setValues(range.getValues());
      range.clearDataValidations();
      sheet.unhideRow(sheet.getRange(1, 1, lastRow, 1));
    }
    if (maxRows > 0) sheet.setRowHeights(1, maxRows, 21);
    sheet.getDrawings().forEach(d => d.remove());
  });

  SpreadsheetApp.flush();
  const url = `https://docs.google.com/spreadsheets/d/${tempFile.getId()}/export?format=xlsx`;
  const htmlOutput = HtmlService.createHtmlOutput(
    `<div style="font-family: Arial, sans-serif; padding: 15px;">
       <h3>Export Ready!</h3>
       <p>Your workbook has been flattened, unhidden, and scrubbed for Excel export.</p>
       <p><a href="${url}" target="_blank" style="background-color: #2e7d32; color: white; padding: 10px 15px; text-decoration: none; border-radius: 4px; display: inline-block;">Download XLSX File</a></p>
       <p style="font-size: 11px; color: #666;">Note: You can safely delete the temporary file from Google Drive after downloading.</p>
     </div>`
  ).setWidth(400).setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Export to Microsoft Excel');
}

function runSystemHealthCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const archive = ss.getSheetByName(ARCHIVE_SHEET);
  const ui = SpreadsheetApp.getUi();
  let report = ["🔍 SYSTEM DIAGNOSTIC REPORT", "--------------------------"];

  if (!inv) {
    ui.alert("CRITICAL ERROR: Inventory Sheet not found.");
    return;
  }

  const headers = inv.getRange(1, 1, 1, inv.getLastColumn()).getValues()[0];
  report.push(headers[ACQ_COL] && headers[ACQ_COL].toString().toLowerCase().includes("acq")
    ? "✅ Column I (Acquisition): Aligned."
    : `⚠️ Column I Mismatch: Found "${headers[ACQ_COL]}"`);

  report.push(headers[FREQ_COL] && headers[FREQ_COL].toString().toLowerCase().includes("repot")
    ? "✅ Column T (Repot Every): Aligned."
    : `⚠️ Column T Mismatch: Found "${headers[FREQ_COL]}"`);

  const data = inv.getDataRange().getValues();
  const seenIDs = new Set();
  let hiddenCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    if (inv.isRowHiddenByUser(row)) {
      hiddenCount++;
      continue;
    }
    const id = data[i][ID_COL];
    const name = data[i][NAME_COL];
    if (!id || !name) continue;

    if (seenIDs.has(id)) report.push(`⚠️ Row ${row}: Duplicate ID ${id}`);
    seenIDs.add(id);

    const sheet = ss.getSheetByName(String(id).trim());
    if (!sheet) report.push(`⚠️ ID ${id}: Missing individual sheet`);
    else if (sheet.isSheetHidden()) report.push(`⚠️ ID ${id}: Sheet is hidden while active`);
  }

  const archiveCount = archive ? Math.max(0, archive.getLastRow() - 1) : 0;
  report.push(`📊 Active Specimens: ${seenIDs.size}`);
  report.push(`📊 Archived Specimens: ${archiveCount} (Hidden Inventory rows: ${hiddenCount})`);

  ui.alert(report.join("\n"));
}
