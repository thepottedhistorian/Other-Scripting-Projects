/**
 * =========================================================================
 * PROJECT: Zöe's Daily Care Tracker - Remote Form Routing Engine
 * PURPOSE: Maps live form submissions directly to the tracking spreadsheet.
 * AUTHOR: Rebecca Short
 * VERSION: 6.0 (Aligned to Row 7 & Verified Physical Column Architecture)
 * =========================================================================
 */

function routeFormToTrackerDirect(e) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const spreadsheetId = scriptProperties.getProperty('SPREADSHEET_ID') || "YOUR_SPREADSHEET_ID_HERE";
    
    if (!spreadsheetId || spreadsheetId === "YOUR_SPREADSHEET_ID_HERE") {
      throw new Error("Missing SPREADSHEET_ID in Script Properties.");
    }

    const ss = SpreadsheetApp.openById(spreadsheetId);
    const trackerSheet = ss.getSheetByName("Tracking"); 
    const darbepoetinSheet = ss.getSheetByName("Darbepoetin Log") || ss.getSheetByName("Darbopoetin Log");
    
    const itemResponses = e.response.getItemResponses();
    const timestampString = e.response.getTimestamp();
    const formDate = new Date(timestampString);
    const timeZone = ss.getSpreadsheetTimeZone();
    const cleanTimeStr = Utilities.formatDate(formDate, timeZone, "h:mm a"); 

    // 1. DATA EXTRACTION
    let responses = {};
    for (let i = 0; i < itemResponses.length; i++) {
      let title = itemResponses[i].getItem().getTitle().trim();
      let response = itemResponses[i].getResponse();
      let stringValue = Array.isArray(response) ? response.join(", ") : response.toString();
      
      if (responses[title]) {
        responses[title] = responses[title] + ", " + stringValue;
      } else {
        responses[title] = stringValue;
      }
    }

    // 2. LOCATE TARGET ROW (Never insert directly; find matching row or top Row 7)
    let targetTrackerRow = findDateRowExact(trackerSheet, formDate);
    
    // If not found, use top data row (Row 7)
    if (targetTrackerRow === -1) {
      targetTrackerRow = 7;
    }
    
    // -------------------------------------------------------------------------
    // 3. EXACT COLUMN MAPPING (Physical Columns)
    // -------------------------------------------------------------------------
    
    // Col B (2): Sub-Q Fluids (ml)
    if (responses['Sub Q Fluids (ml)']) {
      trackerSheet.getRange(targetTrackerRow, 2).setValue(responses['Sub Q Fluids (ml)']);
    }
    
    // Col C (3): Slurry (ml)
    if (responses['Slurry (ml)']) {
      trackerSheet.getRange(targetTrackerRow, 3).setValue(responses['Slurry (ml)']);
    }
    
    // Col G (7): Potassium Gel
    if (responses['Potassium Gel']) {
      trackerSheet.getRange(targetTrackerRow, 7).setValue(responses['Potassium Gel']);
    }
    
    // Col H (8): Fish Oil (ml)
    if (responses['Fish Oil (ml)']) {
      trackerSheet.getRange(targetTrackerRow, 8).setValue(responses['Fish Oil (ml)']);
    }
    
    // Col I (9): Kidney Support Gold (Checkbox)
    if (responses['Kidney Support Gold (ml)']) {
      let val = responses['Kidney Support Gold (ml)'];
      if (val && val.length > 0 && !val.toLowerCase().includes("not given")) {
        trackerSheet.getRange(targetTrackerRow, 9).setValue(true);
      }
    }

    // Col L (12): Darbepoetin (Checkbox & Log Sheet Sync)
    for (let qTitle in responses) {
      let ans = responses[qTitle];
      if (/darb[eo]poetin/i.test(qTitle) || /darb[eo]poetin/i.test(ans)) {
        if (/given/i.test(ans) && !/not given/i.test(ans)) {
          trackerSheet.getRange(targetTrackerRow, 12).setValue(true);
          
          if (darbepoetinSheet) {
            const targetDarbRow = findDateRowExact(darbepoetinSheet, formDate);
            if (targetDarbRow !== -1) {
              const dLogCell = darbepoetinSheet.getRange(targetDarbRow, 2);
              dLogCell.setNumberFormat("@");
              dLogCell.setValue(cleanTimeStr);
              
              const phaseCell = darbepoetinSheet.getRange(targetDarbRow, 3);
              phaseCell.setValue("Dose Given Today").setFontStyle("italic");

              const prevDoses = Number(darbepoetinSheet.getRange(targetDarbRow + 1, 5).getValue()) || 0;
              darbepoetinSheet.getRange(targetDarbRow, 5).setValue(prevDoses + 1);
            }
          }
        }
        break;
      }
    }

    // Col M (13): Amlodipine (Dropdown)
    if (responses['Amlodipine']) {
      let amlVal = responses['Amlodipine'];
      if (amlVal && amlVal.trim().length > 0) {
        updateMultiSelectCell(trackerSheet, targetTrackerRow, 13, amlVal);
      }
    }

    // Col O (15): Other Medications Given (Dropdown)
    if (responses['Other Meds Given?']) {
      let otherVal = responses['Other Meds Given?'];
      if (otherVal && otherVal.trim().length > 0) {
        updateMultiSelectCell(trackerSheet, targetTrackerRow, 15, otherVal);
      }
    }

    // Col P (16): Notes
    if (responses['Notes?']) {
      trackerSheet.getRange(targetTrackerRow, 16).setValue(responses['Notes?']);
    }

    // Col N (14): Total Intake Formula: Sub-Q (B) + Slurry (C) + Fountain (D)
    trackerSheet.getRange(targetTrackerRow, 14).setFormula(
      '=SUM(B' + targetTrackerRow + ':C' + targetTrackerRow + ', D' + targetTrackerRow + ')'
    );

  } catch (error) {
    Logger.log("CRITICAL RUNTIME ERROR: " + error.toString());
  }
}

// =========================================================================
// GLOBAL HELPERS
// =========================================================================

/**
 * Robust date search comparing year, month, and day
 */
function findDateRowExact(sheet, targetDate) {
  if (!sheet) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  
  const t = new Date(targetDate);
  const targetY = t.getFullYear();
  const targetM = t.getMonth();
  const targetD = t.getDate();

  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const cell = values[i][0];
    if (cell) {
      const d = new Date(cell);
      if (!isNaN(d.getTime())) {
        if (d.getFullYear() === targetY && d.getMonth() === targetM && d.getDate() === targetD) {
          return i + 1;
        }
      }
    }
  }
  return -1;
}

function updateMultiSelectCell(sheet, row, col, newValuesStr) {
  if (!newValuesStr) return;
  const cell = sheet.getRange(row, col);
  const existingValue = cell.getValue().toString();
  
  if (!existingValue) {
    cell.setValue(newValuesStr);
  } else {
    let existingItems = existingValue.split(",").map(item => item.trim());
    let incomingItems = newValuesStr.split(",").map(item => item.trim());
    
    incomingItems.forEach(item => {
      if (existingItems.indexOf(item) === -1) {
        existingItems.push(item);
      }
    });
    
    cell.setValue(existingItems.join(", "));
  }
}
