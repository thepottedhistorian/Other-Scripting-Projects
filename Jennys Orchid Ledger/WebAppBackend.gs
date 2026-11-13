/**
 * =============================================================================
 * JENNY'S ORCHID LEDGER - CONTAINER-BOUND WEB APP BACKEND (WebAppBackend.gs)
 * =============================================================================
 * Serves Index.html and bridges the standalone web app UI with Jenny's active
 * workbook, compact A:N specimen tabs, Archived Orchids, and Google Drive.
 * =============================================================================
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle("Jenny's Potted Ledger – Botanical Archivist")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWebAppSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function parseWebAppDate_(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  const s = String(val).trim();
  if (!s || s === '—' || s === '-') return null;

  const ymd = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (ymd) {
    return new Date(parseInt(ymd[1], 10), parseInt(ymd[2], 10) - 1, parseInt(ymd[3], 10));
  }
  const mdy = s.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (mdy) {
    return new Date(parseInt(mdy[3], 10), parseInt(mdy[1], 10) - 1, parseInt(mdy[2], 10));
  }
  const parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function formatWebAppDuration_(startDate, endDate) {
  const start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
  if (end < start) return '0 days';

  let months = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth());
  let days = end.getDate() - start.getDate();

  if (days < 0) {
    months--;
    days += new Date(end.getFullYear(), end.getMonth(), 0).getDate();
  }

  const parts = [];
  if (months > 0) parts.push(`${months} ${months === 1 ? 'month' : 'months'}`);
  if (days > 0 || parts.length === 0) parts.push(`${days} ${days === 1 ? 'day' : 'days'}`);
  return parts.join(' ');
}

function getWebAppLastCompletedBloomMap_(ss) {
  const map = {};
  const idPattern = /^\d+(-\d+)?$/;

  ss.getSheets().forEach(sheet => {
    const name = sheet.getName().trim();
    if (!idPattern.test(name) || sheet.getLastRow() < 20) return;

    const numRows = Math.min(sheet.getLastRow() - 19, 40);
    const bloomCols = sheet.getRange(20, 1, numRows, 3).getValues();
    let latestClosedDuration = '';

    for (let r = 0; r < bloomCols.length; r++) {
      const inVal = bloomCols[r][0];
      const outVal = bloomCols[r][1];
      const durVal = bloomCols[r][2];

      const inStr = String(inVal || '').trim();
      const outStr = String(outVal || '').trim();
      const durStr = String(durVal || '').trim();

      if (inStr && outStr && outStr !== '-' && outStr !== '—') {
        if (durStr && durStr !== '—' && durStr !== '-') {
          latestClosedDuration = durStr;
        } else {
          const d1 = parseWebAppDate_(inVal);
          const d2 = parseWebAppDate_(outVal);
          if (d1 && d2) latestClosedDuration = formatWebAppDuration_(d1, d2);
        }
      }
    }

    if (latestClosedDuration) map[name] = latestClosedDuration;
  });

  return map;
}

/**
 * Fetch all active inventory records from 'Inventory'
 */
function getInventoryData() {
  try {
    const ss = getWebAppSpreadsheet_();
    const sheet = ss.getSheetByName(INVENTORY_SHEET) || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const lastBloomMap = getWebAppLastCompletedBloomMap_(ss);
    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1);

    return rows
      .filter((row, idx) => !sheet.isRowHiddenByUser(idx + 2) && row[0] !== '' && row[0] !== null && row[0] !== undefined)
      .map((row, index) => {
        let record = { _rowIndex: index + 2 };
        headers.forEach((header, colIdx) => {
          let val = row[colIdx];
          if (val instanceof Date && !isNaN(val.getTime())) {
            val = Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
          }
          record[header] = (val !== null && val !== undefined) ? String(val).trim() : '';
        });
        const idKey = String(record['ID']).trim();
        record['_lastCompletedBloomDuration'] = lastBloomMap[idKey] || '';
        return record;
      });
  } catch (err) {
    throw new Error('Failed to fetch inventory: ' + err.message);
  }
}

/**
 * Fetch all archived inventory records from 'Archived Orchids'
 */
function getArchivedInventoryData() {
  try {
    const ss = getWebAppSpreadsheet_();
    const sheet = ss.getSheetByName(ARCHIVE_SHEET);
    if (!sheet) return [];

    const tz = ss.getSpreadsheetTimeZone();
    const lastBloomMap = getWebAppLastCompletedBloomMap_(ss);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0].map(h => String(h).trim());
    const rows = data.slice(1);

    return rows
      .filter(row => row[0] !== '' && row[0] !== null && row[0] !== undefined)
      .map((row, index) => {
        let record = { _rowIndex: index + 2, _isArchived: true };
        headers.forEach((header, colIdx) => {
          let val = row[colIdx];
          if (val instanceof Date && !isNaN(val.getTime())) {
            val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
          }
          record[header] = (val !== null && val !== undefined) ? String(val).trim() : '';
        });

        record['ID'] = record['Orchid ID'] || record['ID'] || '';
        record['Orchid Name'] = record['Name'] || record['Orchid Name'] || '';
        record['Species/Hybrid Lineage'] = record['Species / Hybrid'] || record['Species/Hybrid Lineage'] || '';
        record['Acquisition date'] = record['Acquisition Date'] || record['Acquisition date'] || '';
        record['In Bloom'] = record['Final Status'] || 'Archived';

        const idKey = String(record['ID']).trim();
        record['_lastCompletedBloomDuration'] = lastBloomMap[idKey] || '';

        if (!record['Age'] && record['Acquisition date'] && record['Archived Date']) {
          const acqD = parseWebAppDate_(record['Acquisition date']);
          const archD = parseWebAppDate_(record['Archived Date']);
          if (acqD && archD) record['Age'] = formatWebAppDuration_(acqD, archD);
        }

        return record;
      });
  } catch (err) {
    throw new Error('Failed to fetch archived inventory: ' + err.message);
  }
}

/**
 * Web App Archive Endpoint (uses InventoryAndArchive.gs engine for TMP_ recycling)
 */
function webAppArchiveOrchid(id) {
  try {
    const ok = archiveOrchidById_(String(id).trim(), 'Archived', 'Archived via Web App');
    if (!ok) throw new Error(`Specimen ID #${id} could not be archived.`);
    return { success: true, message: `Specimen #${id} moved to Archived Orchids.` };
  } catch (err) {
    throw new Error('Failed to archive specimen: ' + err.message);
  }
}

/**
 * Creates a new orchid accession from the Web App
 */
function createNewOrchid(newRecord) {
  try {
    const ss = getWebAppSpreadsheet_();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(INVENTORY_SHEET) || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idColIdx = headers.indexOf('ID');
    if (idColIdx === -1) throw new Error("Header 'ID' not found.");

    let assignedId = String(newRecord['ID'] || '').trim();
    if (!assignedId) {
      let maxId = 0;
      for (let i = 1; i < data.length; i++) {
        const num = parseInt(data[i][idColIdx], 10);
        if (!isNaN(num) && num > maxId) maxId = num;
      }
      assignedId = String(maxId + 1);
    }
    newRecord['ID'] = assignedId;
    newRecord['_lastCompletedBloomDuration'] = '';

    const newRow = headers.map(header => newRecord.hasOwnProperty(header) ? newRecord[header] : '');
    sheet.appendRow(newRow);
    const appendedRow = sheet.getLastRow();

    // Reuse TMP_ tab if available, otherwise copy Template
    let idSheet = ss.getSheetByName(assignedId);
    if (!idSheet) {
      const tmpSheet = ss.getSheets().find(s => s.getName().toUpperCase().startsWith('TMP_'));
      const templateSheet = ss.getSheetByName(TEMPLATE_SHEET) || ss.getSheetByName('template');
      if (tmpSheet) {
        idSheet = tmpSheet.setName(assignedId);
      } else if (templateSheet) {
        idSheet = templateSheet.copyTo(ss).setName(assignedId);
      }
      if (idSheet) {
        idSheet.showSheet();
        if (!idSheet.getRange('B2').getFormula() && newRecord['Genus']) idSheet.getRange('B2').setValue(newRecord['Genus']);
        if (!idSheet.getRange('B3').getFormula() && newRecord['Species/Hybrid Lineage']) idSheet.getRange('B3').setValue(newRecord['Species/Hybrid Lineage']);
        if (!idSheet.getRange('B4').getFormula() && newRecord['Alliance']) idSheet.getRange('B4').setValue(newRecord['Alliance']);
        if (!idSheet.getRange('B5').getFormula() && newRecord['Orchid Name']) idSheet.getRange('B5').setValue(newRecord['Orchid Name']);
        if (!idSheet.getRange('B6').getFormula() && newRecord['Endemic To']) idSheet.getRange('B6').setValue(newRecord['Endemic To']);
        if (newRecord['In Bloom']) {
          idSheet.getRange('D5').setValue(newRecord['In Bloom']);
          if (typeof applyBloomStatusFormatting === 'function') applyBloomStatusFormatting(idSheet);
        }
        sheet.getRange(appendedRow, 1).setFormula(`=HYPERLINK("#gid=${idSheet.getSheetId()}", "${assignedId}")`);
      }
    }

    // Add to Maintenance Log
    const maintSheet = ss.getSheetByName(MAINT_LOG_SHEET);
    if (maintSheet) {
      const nextLogRow = Math.max(maintSheet.getLastRow() + 1, 3);
      maintSheet.getRange(nextLogRow, 1).setValue(assignedId);
      maintSheet.getRange(nextLogRow, 2).setValue(newRecord['Orchid Name'] || '').setFontStyle('italic');
    }

    // Create Google Drive photo folder and link card K2:N5
    if (typeof ensureSpecimenDrivePhotoFolder_ === 'function') {
      ensureSpecimenDrivePhotoFolder_(assignedId, newRecord['Orchid Name'], idSheet);
    }

    return {
      success: true,
      record: newRecord,
      message: `Created Specimen #${assignedId} (${newRecord['Orchid Name']}) and Drive photo folder.`
    };
  } catch (err) {
    throw new Error('Failed to create orchid: ' + err.message);
  }
}

/**
 * 3-Way Bloom Status Sync from Passport Modal
 */
function updateSpecimenBloomStatus(id, newStatus) {
  try {
    const ss = getWebAppSpreadsheet_();
    const tz = ss.getSpreadsheetTimeZone();
    const today = new Date();
    const todayIso = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
    const todayLong = Utilities.formatDate(today, tz, 'MMMM d, yyyy');

    const targetId = String(id).trim();
    const sLower = String(newStatus || '').toLowerCase();
    const isNowBlooming = sLower.includes('in bloom') && !sLower.includes('not in bloom');

    const invSheet = ss.getSheetByName(INVENTORY_SHEET) || ss.getSheets()[0];
    const invData = invSheet.getDataRange().getValues();
    const headers = invData[0].map(h => String(h).trim());
    const idColIdx = headers.indexOf('ID');
    const bloomColIdx = headers.indexOf('In Bloom');
    const lastBloomColIdx = headers.indexOf('Last Bloom Date');
    const durationColIdx = headers.indexOf('Bloom Duration');

    let invRowNum = -1;
    for (let i = 1; i < invData.length; i++) {
      if (String(invData[i][idColIdx]).trim() === targetId) {
        invRowNum = i + 1;
        invSheet.getRange(invRowNum, bloomColIdx + 1).setValue(newStatus);
        if (isNowBlooming && lastBloomColIdx !== -1) {
          invSheet.getRange(invRowNum, lastBloomColIdx + 1).setValue(todayIso);
        }
        break;
      }
    }

    const idSheet = ss.getSheetByName(targetId);
    let closedDurationText = '';

    if (idSheet) {
      idSheet.getRange('D5').setValue(newStatus);
      if (typeof applyBloomStatusFormatting === 'function') {
        applyBloomStatusFormatting(idSheet);
      }

      const maxRow = Math.max(idSheet.getLastRow(), 25);
      let dataStartRow = 20;
      let lastFilledRow = dataStartRow - 1;
      let openBloomRow = -1;
      let openBloomStartDate = null;

      for (let r = dataStartRow; r <= maxRow; r++) {
        const rowVals = idSheet.getRange(r, 1, 1, 4).getValues()[0];
        const inVal = rowVals[0];
        const outVal = rowVals[1];
        const inStr = String(inVal || '').trim().toLowerCase();
        const outStr = String(outVal || '').trim();

        if (inVal !== '' && inVal !== null && inStr !== 'in bloom' && inStr !== 'date') {
          lastFilledRow = r;
          if (outStr === '' || outStr === 'null') {
            openBloomRow = r;
            openBloomStartDate = parseWebAppDate_(inVal);
          }
        }
      }

      if (isNowBlooming) {
        if (openBloomRow === -1) {
          const insertRow = lastFilledRow + 1;
          idSheet.getRange(insertRow, 1, 1, 4).setValues([[todayLong, '', '', '']]);
          idSheet.getRange(insertRow, 1).setNumberFormat('MMMM d, yyyy');
        }
      } else if (openBloomRow !== -1) {
        closedDurationText = openBloomStartDate ? formatWebAppDuration_(openBloomStartDate, today) : '0 days';
        idSheet.getRange(openBloomRow, 2, 1, 2).setValues([[todayLong, closedDurationText]]);
        idSheet.getRange(openBloomRow, 2).setNumberFormat('MMMM d, yyyy');
        if (invRowNum !== -1 && durationColIdx !== -1) {
          invSheet.getRange(invRowNum, durationColIdx + 1).setValue(closedDurationText);
        }
      }
    }

    return {
      success: true,
      newStatus: newStatus,
      lastBloomDate: isNowBlooming ? todayIso : null,
      closedDuration: closedDurationText || null,
      message: `Updated #${targetId} to "${newStatus}" and synced Bloom Log.`
    };
  } catch (err) {
    throw new Error('Failed to sync bloom status: ' + err.message);
  }
}

/**
 * Loads images from Google Drive "[ID] - [Orchid Name]" folder & syncs K2:N5
 */
function getSpecimenImages(id) {
  const images = [];
  let folderUrl = '';
  const ss = getWebAppSpreadsheet_();
  const targetIdStr = String(id).trim();
  const sheet = ss.getSheetByName(targetIdStr);

  try {
    const invSheet = ss.getSheetByName(INVENTORY_SHEET);
    let orchidName = '';
    if (invSheet) {
      const invData = invSheet.getDataRange().getValues();
      for (let i = 1; i < invData.length; i++) {
        if (String(invData[i][0]).trim() === targetIdStr) {
          orchidName = String(invData[i][1] || '').trim();
          break;
        }
      }
    }

    if (typeof ensureSpecimenDrivePhotoFolder_ === 'function') {
      const res = ensureSpecimenDrivePhotoFolder_(targetIdStr, orchidName, sheet);
      if (res && res.folder) {
        folderUrl = res.folder.getUrl();
        const files = res.folder.getFiles();
        while (files.hasNext()) {
          const file = files.next();
          if (file.getMimeType().startsWith('image/')) {
            images.push({
              id: file.getId(),
              url: `https://lh3.googleusercontent.com/d/${file.getId()}=w1200`,
              thumbUrl: `https://lh3.googleusercontent.com/d/${file.getId()}=w300`,
              name: file.getName(),
              source: 'drive'
            });
          }
        }
      }
    }
  } catch (e) {
    Logger.log('Error reading drive folder photos for ID ' + targetIdStr + ': ' + e.message);
  }

  return { images: images, folderUrl: folderUrl };
}

/**
 * Reads tabbed logs from compact A:N layout (Bloom A:D, Notes F:G, Repot I:J, Water L:N)
 */
function getOrchidIdSheetLogs(id) {
  try {
    const ss = getWebAppSpreadsheet_();
    const sheet = ss.getSheetByName(String(id).trim());
    if (!sheet || sheet.getLastRow() < 19) {
      return { notes: [], watering: [], repotting: [], bloom: [] };
    }

    const tz = ss.getSpreadsheetTimeZone();
    const formatDate = (val) => {
      if (!val) return '';
      if (val instanceof Date && !isNaN(val.getTime())) {
        return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
      }
      const parsed = parseWebAppDate_(val);
      return parsed ? Utilities.formatDate(parsed, tz, 'yyyy-MM-dd') : String(val).trim();
    };

    const lastCol = sheet.getLastColumn();
    const numDataRows = sheet.getLastRow() - 18;
    const fullGrid = sheet.getRange(19, 1, numDataRows, lastCol).getValues();
    const logs = { notes: [], watering: [], repotting: [], bloom: [] };

    for (let r = 0; r < fullGrid.length; r++) {
      const row = fullGrid[r];
      const sheetRowIndex = 19 + r;

      // Bloom Log (Col A = 1)
      const bDate = row[0];
      const bDateStr = String(bDate || '').trim().toLowerCase();
      if (bDate && bDateStr !== '' && bDateStr !== 'in bloom' && bDateStr !== 'date' && bDateStr !== 'bloom log' && bDateStr !== 'start date') {
        const startDate = parseWebAppDate_(bDate);
        const endDateVal = row[1];
        const endDate = parseWebAppDate_(endDateVal);
        let durStr = row[2] ? String(row[2]).trim() : '';

        if ((!durStr || durStr === '—') && startDate) {
          durStr = formatWebAppDuration_(startDate, endDate || new Date());
        }

        logs.bloom.push({
          rowIndex: sheetRowIndex,
          colIndex: 1,
          inBloom: formatDate(bDate),
          outBloom: formatDate(endDateVal),
          duration: durStr || '—',
          notes: row[3] ? String(row[3]).trim() : ''
        });
      }

      // Notes (Col F = 6)
      const nDate = row[5];
      const nDateStr = String(nDate || '').trim().toLowerCase();
      if (nDate && nDateStr !== '' && nDateStr !== 'date') {
        logs.notes.push({
          rowIndex: sheetRowIndex,
          colIndex: 6,
          date: formatDate(nDate),
          notes: row[6] ? String(row[6]).trim() : ''
        });
      }

      // Repotting Log (Col I = 9)
      const repDate = row[8];
      const repDateStr = String(repDate || '').trim().toLowerCase();
      if (repDate && repDateStr !== '' && repDateStr !== 'date') {
        logs.repotting.push({
          rowIndex: sheetRowIndex,
          colIndex: 9,
          date: formatDate(repDate),
          notes: row[9] ? String(row[9]).trim() : ''
        });
      }

      // Watering Log (Col L = 12)
      const wDate = row[11];
      const wDateStr = String(wDate || '').trim().toLowerCase();
      if (wDate && wDateStr !== '' && wDateStr !== 'date') {
        logs.watering.push({
          rowIndex: sheetRowIndex,
          colIndex: 12,
          date: formatDate(wDate),
          method: row[12] ? String(row[12]).trim() : 'Soak',
          notes: row[13] ? String(row[13]).trim() : ''
        });
      }
    }

    return logs;
  } catch (err) {
    return { notes: [], watering: [], repotting: [], bloom: [], error: err.message };
  }
}

function updateIdSheetLogEntry(id, logType, rowIndex, colIndex, payload) {
  try {
    const ss = getWebAppSpreadsheet_();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(String(id).trim());
    if (!sheet) throw new Error(`Sheet #${id} not found.`);

    const r = Number(rowIndex);
    const c = Number(colIndex);
    const formatLongOrIso = (isoStr) => {
      const d = parseWebAppDate_(isoStr);
      return d ? Utilities.formatDate(d, tz, 'MMMM d, yyyy') : isoStr;
    };

    if (logType === 'notes' || logType === 'repotting') {
      sheet.getRange(r, c, 1, 2).setValues([[formatLongOrIso(payload.date), payload.notes || '']]);
      sheet.getRange(r, c).setNumberFormat('MMMM d, yyyy');
    } else if (logType === 'watering') {
      sheet.getRange(r, c, 1, 3).setValues([[formatLongOrIso(payload.date), payload.method || 'Soak', payload.notes || '']]);
      sheet.getRange(r, c).setNumberFormat('MMMM d, yyyy');
    } else if (logType === 'bloom') {
      const startD = parseWebAppDate_(payload.inBloom);
      const endD = parseWebAppDate_(payload.outBloom);
      const durationText = (startD && endD) ? formatWebAppDuration_(startD, endD) : '';
      const inFormatted = startD ? Utilities.formatDate(startD, tz, 'MMMM d, yyyy') : '';
      const outFormatted = endD ? Utilities.formatDate(endD, tz, 'MMMM d, yyyy') : '';

      sheet.getRange(r, c, 1, 4).setValues([[inFormatted, outFormatted, durationText, payload.notes || '']]);
      sheet.getRange(r, c, 1, 2).setNumberFormat('MMMM d, yyyy');
    }

    return { success: true, message: 'Chronicle entry updated.' };
  } catch (err) {
    throw new Error('Failed to update entry: ' + err.message);
  }
}

function deleteIdSheetLogEntry(id, logType, rowIndex, colIndex) {
  try {
    const ss = getWebAppSpreadsheet_();
    const sheet = ss.getSheetByName(String(id).trim());
    if (!sheet) throw new Error(`Sheet #${id} not found.`);

    const width = (logType === 'bloom') ? 4 : (logType === 'watering' ? 3 : 2);
    sheet.getRange(Number(rowIndex), Number(colIndex), 1, width).deleteCells(SpreadsheetApp.Dimension.ROWS);
    return { success: true, message: 'Chronicle entry removed.' };
  } catch (err) {
    throw new Error('Failed to delete entry: ' + err.message);
  }
}

function logCareToIdSheet(id, actionType, methodOrMedia, noteText) {
  try {
    const ss = getWebAppSpreadsheet_();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(String(id).trim());
    if (!sheet) throw new Error(`Sheet #${id} not found.`);

    const todayLong = Utilities.formatDate(new Date(), tz, 'MMMM d, yyyy');
    let targetCol = 6; // Observation (Col F)
    if (actionType === 'Watering') targetCol = 12;     // Col L
    else if (actionType === 'Repotting') targetCol = 9;// Col I

    let targetRow = 20;
    const checkRows = Math.max(sheet.getLastRow() - 19, 1);
    const colValues = sheet.getRange(20, targetCol, checkRows, 1).getValues();
    for (let i = 0; i < colValues.length; i++) {
      if (String(colValues[i][0] || '').trim() !== '') {
        targetRow = 20 + i + 1;
      }
    }

    if (actionType === 'Watering') {
      sheet.getRange(targetRow, targetCol, 1, 3).setValues([[todayLong, methodOrMedia || 'Soak', noteText || '']]);
    } else if (actionType === 'Repotting') {
      sheet.getRange(targetRow, targetCol, 1, 2).setValues([[todayLong, noteText || methodOrMedia || 'Fresh mix']]);
    } else {
      sheet.getRange(targetRow, targetCol, 1, 2).setValues([[todayLong, noteText || '']]);
    }
    sheet.getRange(targetRow, targetCol).setNumberFormat('MMMM d, yyyy');

    return { success: true, message: `Logged ${actionType} on Sheet #${id}` };
  } catch (err) {
    throw new Error('Failed to append to ID sheet: ' + err.message);
  }
}

function updateOrchidRecord(updatedRecord) {
  try {
    const ss = getWebAppSpreadsheet_();
    const sheet = ss.getSheetByName(INVENTORY_SHEET) || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idColIdx = headers.indexOf('ID');
    const targetId = String(updatedRecord['ID']).trim();
    let targetRowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idColIdx]).trim() === targetId) {
        targetRowIndex = i + 1;
        break;
      }
    }
    if (targetRowIndex === -1) throw new Error(`Specimen ID #${targetId} not found.`);

    headers.forEach((header, colIdx) => {
      if (updatedRecord.hasOwnProperty(header)) {
        sheet.getRange(targetRowIndex, colIdx + 1).setValue(updatedRecord[header]);
      }
    });

    if (updatedRecord['Orchid Name'] && typeof ensureSpecimenDrivePhotoFolder_ === 'function') {
      ensureSpecimenDrivePhotoFolder_(targetId, updatedRecord['Orchid Name'], ss.getSheetByName(targetId));
    }

    return { success: true, message: `Specimen #${targetId} updated successfully.` };
  } catch (err) {
    throw new Error('Update failed: ' + err.message);
  }
}

function executeBatchAction(specimenIds, actionType, extraFields) {
  try {
    const ss = getWebAppSpreadsheet_();
    const sheet = ss.getSheetByName(INVENTORY_SHEET) || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const idColIdx = headers.indexOf('ID');

    const idSet = new Set(specimenIds.map(id => String(id).trim()));
    let updatedCount = 0;

    for (let i = 1; i < data.length; i++) {
      const currentId = String(data[i][idColIdx]).trim();
      if (idSet.has(currentId)) {
        const rowNumber = i + 1;
        if (extraFields && typeof extraFields === 'object') {
          headers.forEach((header, colIdx) => {
            if (extraFields.hasOwnProperty(header)) {
              sheet.getRange(rowNumber, colIdx + 1).setValue(extraFields[header]);
            }
          });
        }
        updatedCount++;
      }
    }

    return { success: true, updatedCount: updatedCount, message: `Applied ${actionType} to ${updatedCount} specimens.` };
  } catch (err) {
    throw new Error('Batch operation failed: ' + err.message);
  }
}
