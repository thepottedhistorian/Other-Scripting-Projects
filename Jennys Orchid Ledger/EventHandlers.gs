/**
 * =============================================================================
 * JENNY'S ORCHID LEDGER - CENTRAL EVENT HANDLER (EventHandlers.gs)
 * =============================================================================
 * Purpose: Triggers on sheet edits to sync Maintenance Log entries, Inventory
 *          bloom status changes, and individual ID sheet logs (Compact A:N).
 * =============================================================================
 */

function onEdit(e) {
  if (!e) return;
  const ss = e.source;
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName().trim();

  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const startCol = range.getColumn();
  const lastCol = range.getLastColumn();

  // 1. MAINTENANCE LOG SYNC
  if (sheetName === MAINT_LOG_SHEET) {
    const ML_ID_COL = 1;
    const ML_METHOD_COL = 3;
    const ML_DATE_COL = 4;
    const ML_NOTES_COL = 5;
    const ML_DUE_COL = 7;
    const ML_REPOT_DATE = 8;

    for (let i = 0; i < numRows; i++) {
      const currentRow = startRow + i;
      if (currentRow <= 2) continue; // Protects Row 1 banner & Row 2 headers (e.g., D2)

      const orchidID = sheet.getRange(currentRow, ML_ID_COL).getValue();
      if (!orchidID) continue;
      const orchidSheet = ss.getSheetByName(String(orchidID).trim());

      // --- WATERING LOGIC (Cols L:N) ---
      if (startCol <= ML_NOTES_COL && lastCol >= ML_METHOD_COL) {
        const methodVal = sheet.getRange(currentRow, ML_METHOD_COL).getValue();
        const notesVal = sheet.getRange(currentRow, ML_NOTES_COL).getValue();

        if (methodVal) {
          const now = new Date();
          sheet.getRange(currentRow, ML_DATE_COL).setValue(now);

          if (orchidSheet) {
            let targetRow = orchidSheet.getRange("L100").getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
            if (targetRow < 20) targetRow = 20;

            orchidSheet.getRange(targetRow, 12).setValue(now).setNumberFormat('MMMM d, yyyy');
            orchidSheet.getRange(targetRow, 13).setValue(methodVal);
            orchidSheet.getRange(targetRow, 14).setValue(notesVal || "");
          }
        }
      }

      // --- REPOTTING LOGIC (Cols I:J) ---
      if (startCol <= 9 && lastCol >= ML_REPOT_DATE) {
        const repotDateVal = sheet.getRange(currentRow, ML_REPOT_DATE).getValue();
        const repotNotes = sheet.getRange(currentRow, 9).getValue();
        let repotDateObj = (repotDateVal instanceof Date && !isNaN(repotDateVal.getTime())) ? repotDateVal : (repotDateVal ? new Date(repotDateVal) : null);

        if (repotDateObj && !isNaN(repotDateObj.getTime()) && orchidSheet) {
          let targetRow = orchidSheet.getRange("I100").getNextDataCell(SpreadsheetApp.Direction.UP).getRow() + 1;
          if (targetRow < 20) targetRow = 20;

          orchidSheet.getRange(targetRow, 9).setValue(repotDateObj).setNumberFormat('MMMM d, yyyy');
          if (repotNotes) orchidSheet.getRange(targetRow, 10).setValue(repotNotes);

          const invSheet = ss.getSheetByName(INVENTORY_SHEET);
          const invData = invSheet.getDataRange().getValues();

          for (let j = 1; j < invData.length; j++) {
            if (String(invData[j][ID_COL]).trim() === String(orchidID).trim()) {
              const freqRaw = invData[j][FREQ_COL];
              const parsedMonths = (typeof parseFrequencyToMonths === 'function') ? parseFrequencyToMonths(freqRaw) : null;
              const months = (typeof parsedMonths === 'number' && parsedMonths > 0) ? parsedMonths : 12;

              const nextDate = new Date(repotDateObj);
              nextDate.setMonth(nextDate.getMonth() + months);

              sheet.getRange(currentRow, ML_DUE_COL).setValue(nextDate).setNumberFormat('MMMM d, yyyy');
              invSheet.getRange(j + 1, NEXT_COL + 1).setValue(nextDate).setNumberFormat('MMMM d, yyyy');

              const d10 = orchidSheet.getRange("D10");
              d10.setValue(nextDate).setNumberFormat('MMMM d, yyyy').setHorizontalAlignment("center");
              break;
            }
          }
        }
      }
    }
  }

  // 2. INVENTORY MASTER SYNC (Bloom Status Col K -> D5 & Bloom Log Cols A:C)
  if (sheetName === INVENTORY_SHEET) {
    for (let i = 0; i < numRows; i++) {
      const currentRow = startRow + i;
      if (currentRow < 2) continue;

      const orchidID = sheet.getRange(currentRow, ID_COL + 1).getValue();
      if (!orchidID) continue;
      const orchidSheet = ss.getSheetByName(String(orchidID).trim());

      if (startCol === BLOOM_STATUS_COL + 1 && orchidSheet) {
        const rawValue = range.getValue().toString();
        orchidSheet.getRange("D5").setValue(rawValue);

        const cleanVal = rawValue.replace(/[^\w\s]/gi, '').trim().toLowerCase();
        const now = new Date();

        if (cleanVal === "in bloom") {
          sheet.getRange(currentRow, LAST_BLOOM_COL + 1).setValue(now);
          logBloomEvent_(orchidSheet, "start", now);
        } else if (cleanVal === "not in bloom") {
          const inventoryBloomDate = sheet.getRange(currentRow, LAST_BLOOM_COL + 1).getValue();
          const targetEndDate = (inventoryBloomDate instanceof Date && !isNaN(inventoryBloomDate.getTime())) ? inventoryBloomDate : now;
          logBloomEvent_(orchidSheet, "end", targetEndDate);
        }

        if (typeof applyBloomStatusFormatting === 'function') {
          applyBloomStatusFormatting(orchidSheet);
        }
      }
    }
  }

  // 3. INDIVIDUAL ID SHEET BLOOM DURATION RECALC (Col B edit on Row 20+)
  if (/^\d+(-\d+)?$/.test(sheetName) && startCol === 2 && startRow >= 20) {
    if (typeof updateSheetBloomDurations === 'function') {
      updateSheetBloomDurations(sheet);
    }
  }
}

function logBloomEvent_(orchidSheet, type, targetDate) {
  const logRange = orchidSheet.getRange("A20:C45");
  const logData = logRange.getValues();
  let openRowIndices = [];
  let emptyRowIndex = -1;

  for (let i = 0; i < logData.length; i++) {
    const hasStart = logData[i][0] !== "" && logData[i][0] !== null;
    const hasEnd = logData[i][1] !== "" && logData[i][1] !== null;

    if (hasStart && !hasEnd) openRowIndices.push(i);
    if (!hasStart && !hasEnd && emptyRowIndex === -1) emptyRowIndex = i;
  }

  const dateToApply = (targetDate instanceof Date && !isNaN(targetDate.getTime())) ? targetDate : new Date();

  if (type === "start") {
    const lastRowIndex = emptyRowIndex > 0 ? (emptyRowIndex - 1) : (logData.length - 1);
    const existingDate = logData[lastRowIndex][0];
    if (existingDate instanceof Date && existingDate.toDateString() === dateToApply.toDateString() && logData[lastRowIndex][1] === "") {
      return;
    }
    if (emptyRowIndex !== -1) {
      orchidSheet.getRange(20 + emptyRowIndex, 1).setValue(dateToApply).setNumberFormat('MMMM d, yyyy');
    }
  } else if (type === "end" && openRowIndices.length > 0) {
    openRowIndices.forEach(idx => {
      const startDate = logData[idx][0] instanceof Date ? logData[idx][0] : new Date(logData[idx][0]);
      orchidSheet.getRange(20 + idx, 2).setValue(dateToApply).setNumberFormat('MMMM d, yyyy');
      const durationStr = (typeof calculateDurationString_ === 'function') ? calculateDurationString_(startDate, dateToApply) : "";
      if (durationStr) orchidSheet.getRange(20 + idx, 3).setValue(durationStr);
    });
  }
}
