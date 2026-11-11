/**
 * =========================================================================
 * WATCHER & DASHBOARD LAUNCHER: Zöe's Darbepoetin Care Log Router
 * AUTHOR: Rebecca Short
 * VERSION: 10.0 (Dedicated Darbepoetin Checkbox Handler)
 * =========================================================================
 */

/**
 * WATCHER: Listens for Column L (12) checkbox clicks on "Tracking".
 * Stamps time and increments dose counter on "Darbepoetin Log".
 */
function onEdit(e) {
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Core Filter: Must be on "Tracking" sheet, Column L (12), and checked true
  const col = range.getColumn();
  if (sheet.getName() !== "Tracking" || col !== 12 || range.getValue() !== true) {
    return;
  }

  // Prevent double-logging on rapid clicks
  const lock = LockService.getScriptLock();
  try { 
    lock.waitLock(5000); 
  } catch (err) { 
    return; 
  }
  
  if (e.oldValue === "TRUE") { 
    lock.releaseLock(); 
    return; 
  }

  const row = range.getRow();
  const targetDateStr = sheet.getRange(row, 1).getDisplayValue();
  const timeStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "h:mm a");

  const logSheet = ss.getSheetByName("Darbepoetin Log") || ss.getSheetByName("Darbopoetin Log");
  if (!logSheet) { 
    lock.releaseLock(); 
    return; 
  }
  
  const data = logSheet.getRange("A:A").getDisplayValues();
  let targetRow = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === targetDateStr) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow > 0) {
    const cell = logSheet.getRange(targetRow, 2); // Column B on log sheet
    cell.setNumberFormat("@"); 
    cell.setValue(timeStr);
    cell.setFontWeight("normal");
    cell.setHorizontalAlignment("left");

    // Increment dose count from row below if applicable
    const prevDoses = Number(logSheet.getRange(targetRow + 1, 5).getValue()) || 0;
    logSheet.getRange(targetRow, 5).setValue(prevDoses + 1);

    ss.toast("Darbepoetin logged at " + timeStr, "Success");
  } else {
    ss.toast("Error: Looking for '" + targetDateStr + "' on Darbepoetin Log.", "Sync Error");
  }

  lock.releaseLock();
}

/**
 * OPENS THE LIVE PROTOCOL DASHBOARD IN A NEW TAB
 * Attached to the "Current Protocol Dashboard" drawing shape button.
 */
function openProtocolDashboard() {
  const url = "https://script.google.com/macros/s/AKfycbz_Ru7UZrjy-NoGUiugS6K6Q_GPI8LjKqAtB9UeuG_UpVhfMzEbYnQyRhHzkKRM1wJ3OQ/exec";
  const html = HtmlService.createHtmlOutput(
    '<html><script>' +
    'window.open("' + url + '", "_blank");' +
    'google.script.host.close();' +
    '</script><body><p style="font-family:sans-serif;font-size:12px;color:#555;">Opening Dashboard...</p></body></html>'
  ).setWidth(50).setHeight(50);
  
  SpreadsheetApp.getUi().showModalDialog(html, "Opening Dashboard...");
}
