/**
 * =============================================================================
 * JENNY'S ORCHID LEDGER - CARE & REPOTTING TOOLS (CareAndRepotTools.gs)
 * =============================================================================
 * Purpose: Powers the Bulk Watering modal, historical date sync, fertilizer
 *          notes, log deduplication, and batch repot date calculations.
 * =============================================================================
 */

function showWateringModal() {
  const html = HtmlService.createHtmlOutputFromFile('WateringDialog')
    .setWidth(360)
    .setHeight(420);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Log Bulk Watering');
}

function getCurrentRangeString() {
  const selection = SpreadsheetApp.getActiveRangeList();
  if (!selection) return "";

  let rows = [];
  selection.getRanges().forEach(range => {
    const start = range.getRow();
    const end = range.getLastRow();
    for (let r = start; r <= end; r++) {
      if (r >= 3) rows.push(r);
    }
  });

  return [...new Set(rows)].sort((a, b) => a - b).join(", ");
}

function processBulkWatering(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MAINT_LOG_SHEET);
  const now = new Date();

  if (!sheet) throw new Error("Maintenance Log sheet not found.");

  let rows = [];
  const parts = data.range.split(/[\s,]+/);
  parts.forEach(part => {
    if (part.includes(':')) {
      const bounds = part.split(':');
      const start = parseInt(bounds[0], 10);
      const end = parseInt(bounds[1], 10);
      for (let r = start; r <= end; r++) { if (!isNaN(r)) rows.push(r); }
    } else {
      const r = parseInt(part, 10);
      if (!isNaN(r)) rows.push(r);
    }
  });

  rows = [...new Set(rows)].filter(r => r >= 3);
  if (rows.length === 0) throw new Error("No valid rows selected.");

  rows.forEach(currentRow => {
    const orchidID = sheet.getRange(currentRow, 1).getValue();

    sheet.getRange(currentRow, 3).setValue(data.method);
    sheet.getRange(currentRow, 4).setValue(now);
    sheet.getRange(currentRow, 5).setValue(data.notes);

    if (orchidID) {
      const orchidSheet = ss.getSheetByName(String(orchidID).trim());
      if (orchidSheet) {
        let targetRow = orchidSheet.getRange("L100").getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
        if (targetRow < 20) targetRow = 20;

        orchidSheet.getRange(targetRow, 12).setValue(now).setNumberFormat('MMMM d, yyyy');
        orchidSheet.getRange(targetRow, 13).setValue(data.method);
        orchidSheet.getRange(targetRow, 14).setValue(data.notes || "");
      }
    }
  });

  return "Success! Updated " + rows.length + " orchid(s).";
}

function appendFertilizerNotes() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt("Add Fertilizer Note", "Enter fertilizer note to append to each specimen's latest watering:", ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return;

  const noteContent = prompt.getResponseText().trim();
  if (!noteContent) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const idPattern = /^\d+(-\d+)?$/;
  let updateCount = 0;

  sheets.forEach(sheet => {
    if (!idPattern.test(sheet.getName().trim()) || sheet.getLastRow() < 20) return;

    const columnL = sheet.getRange(20, 12, sheet.getLastRow() - 19, 1).getValues();
    let actualLastRow = 0;

    for (let i = columnL.length - 1; i >= 0; i--) {
      if (columnL[i][0] !== "" && columnL[i][0] !== null) {
        actualLastRow = i + 20;
        break;
      }
    }

    if (actualLastRow > 0) {
      const cell = sheet.getRange(actualLastRow, 14);
      const existingNote = String(cell.getValue()).trim();
      if (existingNote.includes(noteContent)) return;

      cell.setValue(existingNote ? `${existingNote}; ${noteContent}` : noteContent);
      updateCount++;
    }
  });

  ui.alert(`Success: Fertilizer note added to ${updateCount} orchid logs.`);
}

function syncMaintenanceByDatePrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Historical Sync', 'Enter the date to sync (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const targetDateStr = response.getResponseText().trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(targetDateStr)) {
    ui.alert("Invalid format. Please use YYYY-MM-DD (e.g., 2026-04-27).");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(MAINT_LOG_SHEET);
  const logData = logSheet.getDataRange().getValues();
  let syncCount = 0;

  for (let i = 2; i < logData.length; i++) {
    const orchidID = logData[i][0];
    const method = logData[i][2];
    const dateTime = logData[i][3];
    const notes = logData[i][4];

    if (dateTime instanceof Date && orchidID) {
      const formattedDate = Utilities.formatDate(dateTime, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
      if (formattedDate === targetDateStr) {
        const orchidSheet = ss.getSheetByName(String(orchidID).trim());
        if (orchidSheet) {
          let targetRow = orchidSheet.getRange("L100").getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
          if (targetRow < 20) targetRow = 20;

          const lastEntry = orchidSheet.getRange(Math.max(1, targetRow - 1), 12).getValue();
          if (lastEntry.toString() === dateTime.toString()) continue;

          orchidSheet.getRange(targetRow, 12).setValue(dateTime).setNumberFormat('MMMM d, yyyy');
          orchidSheet.getRange(targetRow, 13).setValue(method);
          orchidSheet.getRange(targetRow, 14).setValue(notes || "");
          syncCount++;
        }
      }
    }
  }

  ui.alert("Sync Complete. Updated " + syncCount + " records for " + targetDateStr);
}

function cleanDuplicateOrchidEntriesByDate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const idPattern = /^\d+(-\d+)?$/;
  let totalDeleted = 0;

  sheets.forEach(sheet => {
    if (!idPattern.test(sheet.getName().trim())) return;
    const startRow = 20;
    const numRowsToCheck = sheet.getMaxRows() - startRow + 1;
    if (numRowsToCheck < 2) return;

    const range = sheet.getRange(startRow, 12, numRowsToCheck, 3);
    const displayValues = range.getDisplayValues();

    let lastPopulatedIdx = -1;
    for (let i = displayValues.length - 1; i >= 0; i--) {
      if (displayValues[i][0] !== "") {
        lastPopulatedIdx = i;
        break;
      }
    }

    if (lastPopulatedIdx < 1) return;

    for (let i = lastPopulatedIdx; i > 0; i--) {
      if (displayValues[i][0] === displayValues[i - 1][0] && displayValues[i][1] === displayValues[i - 1][1] && displayValues[i][0] !== "") {
        sheet.getRange(startRow + i, 12, 1, 3).deleteCells(SpreadsheetApp.Dimension.ROWS);
        totalDeleted++;
      }
    }
  });

  try {
    SpreadsheetApp.getUi().alert(`Cleanup Complete: Removed ${totalDeleted} duplicate watering entries.`);
  } catch (e) {}
}

/**
 * Repot Date Preview & Batch Recalculation Engine
 */
function showRepotPreviewModal() {
  const htmlTemplate = HtmlService.createTemplateFromFile('repot-preview');
  htmlTemplate.previewData = previewRecalcRepotDates();
  const html = htmlTemplate.evaluate().setWidth(900).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, 'Repot Date Preview');
}

function previewRecalcRepotDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  if (!inv) throw new Error("Inventory sheet not found.");

  const invData = inv.getDataRange().getValues();
  const previewData = [];

  for (let i = 1; i < invData.length; i++) {
    const rowNum = i + 1;
    if (inv.isRowHiddenByUser(rowNum)) continue;
    const id = invData[i][ID_COL];
    if (!id || id.toString().trim() === "" || isNaN(id)) continue;

    const plantName = invData[i][NAME_COL] || "Unknown Orchid";
    const acqRaw = invData[i][ACQ_COL];
    const freqRaw = invData[i][FREQ_COL];
    const freqMonths = parseFrequencyToMonths(freqRaw) || 12;

    const lastRepot = getLastRepotDateForID(id);
    let acqDate = (acqRaw instanceof Date && !isNaN(acqRaw.getTime())) ? acqRaw : (acqRaw ? new Date(acqRaw) : null);
    if (acqDate && isNaN(acqDate.getTime())) acqDate = null;

    const baseUsed = lastRepot ? "Last Repot" : (acqDate ? "Acquisition" : "None");
    const nextDate = calculateNextRepotDate_(acqDate, lastRepot, freqMonths);

    previewData.push({
      row: rowNum,
      id: id,
      name: plantName,
      acquisition: acqDate ? Utilities.formatDate(acqDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : "N/A",
      lastRepot: lastRepot ? Utilities.formatDate(lastRepot, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : "None",
      baseUsed: baseUsed,
      frequency: freqRaw || "1 Year",
      months: freqMonths,
      next: nextDate ? Utilities.formatDate(nextDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") : "N/A",
      status: nextDate ? "OK" : "MISSING BASE DATE"
    });
  }

  return previewData;
}

function recalcAllPlantRepotDueDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const log = ss.getSheetByName(MAINT_LOG_SHEET);
  if (!inv) throw new Error("Inventory sheet not found.");

  const invData = inv.getDataRange().getValues();
  let updatedCount = 0;

  for (let i = 1; i < invData.length; i++) {
    if (inv.isRowHiddenByUser(i + 1)) continue;
    const id = invData[i][ID_COL];
    if (!id || id.toString().trim() === "" || isNaN(id)) continue;

    const acqRaw = invData[i][ACQ_COL];
    const freqRaw = invData[i][FREQ_COL];
    const freqMonths = parseFrequencyToMonths(freqRaw) || 12;

    const lastRepot = getLastRepotDateForID(id);
    let acqDate = (acqRaw instanceof Date && !isNaN(acqRaw.getTime())) ? acqRaw : (acqRaw ? new Date(acqRaw) : null);
    if (acqDate && isNaN(acqDate.getTime())) acqDate = null;

    const nextDate = calculateNextRepotDate_(acqDate, lastRepot, freqMonths);

    if (nextDate) {
      inv.getRange(i + 1, NEXT_COL + 1).setValue(nextDate).setNumberFormat('MMMM d, yyyy');

      if (log) {
        const logVals = log.getDataRange().getValues();
        for (let r = logVals.length - 1; r >= 2; r--) {
          if (String(logVals[r][0]).trim() === String(id).trim()) {
            log.getRange(r + 1, 7).setValue(nextDate).setNumberFormat('MMMM d, yyyy');
            break;
          }
        }
      }

      const orchidSheet = ss.getSheetByName(String(id).trim());
      if (orchidSheet) {
        orchidSheet.getRange("D10").setValue(nextDate).setNumberFormat('MMMM d, yyyy').setHorizontalAlignment("center");
      }
      updatedCount++;
    }
  }

  return updatedCount;
}

function getLastRepotDateForID(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let latest = null;

  const orchidSheet = ss.getSheetByName(String(id).trim());
  if (orchidSheet && orchidSheet.getLastRow() >= 20) {
    const sheetData = orchidSheet.getRange(20, 9, orchidSheet.getLastRow() - 19, 1).getValues();
    for (let j = 0; j < sheetData.length; j++) {
      let d = sheetData[j][0];
      if (d) {
        if (!(d instanceof Date)) d = new Date(d);
        if (!isNaN(d.getTime()) && (!latest || d > latest)) latest = d;
      }
    }
  }

  if (!latest) {
    const log = ss.getSheetByName(MAINT_LOG_SHEET);
    if (log && log.getLastRow() >= 3) {
      const logData = log.getDataRange().getValues();
      for (let k = 2; k < logData.length; k++) {
        if (String(logData[k][0]).trim() === String(id).trim() && logData[k][7]) {
          let d = logData[k][7];
          if (!(d instanceof Date)) d = new Date(d);
          if (!isNaN(d.getTime()) && (!latest || d > latest)) latest = d;
        }
      }
    }
  }
  return latest;
}

function calculateNextRepotDate_(acqDate, lastRepotDate, freqMonths) {
  let base = (lastRepotDate instanceof Date && !isNaN(lastRepotDate.getTime())) ? new Date(lastRepotDate)
    : ((acqDate instanceof Date && !isNaN(acqDate.getTime())) ? new Date(acqDate) : null);

  if (!base || !freqMonths) return null;
  const next = new Date(base);
  next.setMonth(next.getMonth() + freqMonths);
  return next;
}

function parseFrequencyToMonths(freqRaw) {
  if (!freqRaw) return null;
  const text = freqRaw.toString().toLowerCase().trim();
  const matches = text.match(/(\d+(\.\d+)?)/g);
  if (!matches) return null;
  const num = parseFloat(matches[matches.length - 1]);
  return text.includes("year") ? num * 12 : num;
}

function syncRepotDatesFromInventory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName(INVENTORY_SHEET);
  const maintSheet = ss.getSheetByName(MAINT_LOG_SHEET);
  if (!invSheet) throw new Error("Inventory sheet not found.");

  const invData = invSheet.getDataRange().getValues();
  const repotMap = new Map();

  for (let i = 1; i < invData.length; i++) {
    const id = invData[i][ID_COL];
    const nextDue = invData[i][NEXT_COL];
    if (id !== "" && id !== null && nextDue) {
      const dateObj = (nextDue instanceof Date && !isNaN(nextDue.getTime())) ? nextDue : new Date(nextDue);
      if (!isNaN(dateObj.getTime())) repotMap.set(String(id).trim(), dateObj);
    }
  }

  let updatedMaintCount = 0;
  let updatedSheetCount = 0;

  if (maintSheet && maintSheet.getLastRow() >= 3) {
    const maintData = maintSheet.getDataRange().getValues();
    for (let r = 2; r < maintData.length; r++) {
      const maintId = String(maintData[r][0]).trim();
      if (maintId && repotMap.has(maintId)) {
        maintSheet.getRange(r + 1, 7).setValue(repotMap.get(maintId)).setNumberFormat('MMMM d, yyyy');
        updatedMaintCount++;
      }
    }
  }

  repotMap.forEach((targetDate, idStr) => {
    const orchidSheet = ss.getSheetByName(idStr);
    if (orchidSheet) {
      orchidSheet.getRange("D10").setValue(targetDate).setNumberFormat('MMMM d, yyyy').setHorizontalAlignment("center");
      updatedSheetCount++;
    }
  });

  try {
    SpreadsheetApp.getUi().alert(`Repot Sync Complete!\nUpdated ${updatedMaintCount} rows on Maintenance Log and ${updatedSheetCount} Orchid ID sheets.`);
  } catch (e) {}
}
