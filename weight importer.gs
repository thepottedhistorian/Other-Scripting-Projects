/**
 * PROJECT: Zöe's Daily Care Tracker - Weight Log Automation
 * PURPOSE: Parses CSV attachments from Gmail to log weight data into the
 *          'Weight Log' tab and sync latest weight to Column E of 'Tracking'.
 * VERSION: 6.0 (Single-Source Insertion Architecture)
 */

const WT_CONFIG = {
  LABEL_NAME: "Zoes Logs/Zoe_Weight_Data",
  WEIGHT_LOG_SHEET: "Weight Log",
  TRACKING_SHEET: "Tracking",
  MIN_WEIGHT_THRESHOLD: 8.1,
  WEIGHT_COL_TRACKING: 5, // Column E: Weight (lbs)
  LOG_START_ROW: 7,
  COLORS: ["#f3f3f3", "#ffffff"]
};

function importZoeWeightsFromEmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weightSheet = ss.getSheetByName(WT_CONFIG.WEIGHT_LOG_SHEET);
  const trackingSheet = ss.getSheetByName(WT_CONFIG.TRACKING_SHEET);

  const label = GmailApp.getUserLabelByName(WT_CONFIG.LABEL_NAME);
  if (!label) return;

  const threads = label.getThreads().filter(thread => thread.isUnread());
  if (threads.length === 0) return;

  const lastRow = weightSheet.getLastRow();
  let existingKeys = {};
  if (lastRow >= WT_CONFIG.LOG_START_ROW) {
    const data = weightSheet.getRange(WT_CONFIG.LOG_START_ROW, 2, (lastRow - WT_CONFIG.LOG_START_ROW) + 1, 1).getValues();
    data.forEach(row => {
      if (row[0] instanceof Date) {
        let key = (row[0].getMonth() + 1) + "/" + row[0].getDate() + " " + 
                  Utilities.formatDate(row[0], ss.getSpreadsheetTimeZone(), "HH:mm");
        existingKeys[key] = true;
      }
    });
  }

  threads.forEach(thread => {
    const attachments = thread.getMessages().pop().getAttachments();

    attachments.forEach(attachment => {
      if (attachment.getName().toLowerCase().includes("pooper_scooper")) {
        const csvData = Utilities.parseCsv(attachment.getDataAsString());
        let newEntries = [];

        for (let j = 1; j < csvData.length; j++) {
          const activity = csvData[j][0];
          const dateStr = csvData[j][1];
          const weightRaw = csvData[j][2];

          const weight = parseFloat(weightRaw.replace(/[^\d.-]/g, ''));
          
          let cleanDateStr = dateStr;
          if (dateStr && dateStr.includes('at')) {
            cleanDateStr = dateStr.replace(/\s+at\s+/, ' ') + ' 2026';
          }
          
          let properDate = new Date(cleanDateStr);
          if (isNaN(properDate.getTime()) || properDate.getFullYear() < 2020) {
            properDate = new Date(cleanDateStr);
            properDate.setFullYear(2026);
          }

          let lookupKey = (properDate.getMonth() + 1) + "/" + properDate.getDate() + " " + 
                          Utilities.formatDate(properDate, ss.getSpreadsheetTimeZone(), "HH:mm");
          
          if (activity === "Weight recorded" && weight >= WT_CONFIG.MIN_WEIGHT_THRESHOLD && !existingKeys[lookupKey]) {
            newEntries.push(["Weight recorded", properDate, weight]);
            existingKeys[lookupKey] = true;
          }
        }

        if (newEntries.length > 0) {
          newEntries.sort((a, b) => a[1] - b[1]);
          
          for (let n = 0; n < newEntries.length; n++) {
            weightSheet.insertRowBefore(WT_CONFIG.LOG_START_ROW);
            weightSheet.getRange(WT_CONFIG.LOG_START_ROW, 1, 1, 3).setValues([newEntries[n]]);
            weightSheet.getRange(WT_CONFIG.LOG_START_ROW, 3).setHorizontalAlignment("right");
            
            updateTrackingWeight(trackingSheet, newEntries[n][1], newEntries[n][2]);
          }
        }
      }
    });
    thread.markRead();
  });

  manuallyRefreshWeightLog(); 
}

function updateTrackingWeight(sheet, msgDate, weight) {
  if (!sheet) return;

  let row = findTrackingRowForDate(sheet, msgDate);

  if (row === -1) {
    addMidnightRow();
    row = findTrackingRowForDate(sheet, msgDate);
  }

  if (row > 0) {
    sheet.getRange(row, WT_CONFIG.WEIGHT_COL_TRACKING).setValue(weight);
    Logger.log("Set weight " + weight + " at Row " + row);
  }
}

function manuallyRefreshWeightLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weightSheet = ss.getSheetByName(WT_CONFIG.WEIGHT_LOG_SHEET);
  const startRow = WT_CONFIG.LOG_START_ROW;
  if (!weightSheet) return;

  const lastRow = weightSheet.getLastRow();
  if (lastRow < startRow) return;

  const range = weightSheet.getRange(startRow, 1, (lastRow - startRow) + 1, 3);
  const data = range.getValues();
  range.setBackground(null);

  let colorIndex = 0;
  let currentGroupDate = data[0][1] instanceof Date ? data[0][1].toDateString() : "";
  let groupStartRow = 0;

  for (let i = 0; i < data.length; i++) {
    const rowDateObj = data[i][1];
    const rowDateStr = rowDateObj instanceof Date ? rowDateObj.toDateString() : "";

    if (rowDateStr !== currentGroupDate) {
      weightSheet.getRange(startRow + groupStartRow, 1, i - groupStartRow, 3)
                 .setBackground(WT_CONFIG.COLORS[colorIndex % 2]);
      colorIndex++;
      groupStartRow = i;
      currentGroupDate = rowDateStr;
    }

    if (i === data.length - 1) {
      weightSheet.getRange(startRow + groupStartRow, 1, (i - groupStartRow) + 1, 3)
                 .setBackground(WT_CONFIG.COLORS[colorIndex % 2]);
    }
  }
}

/**
 * ROBUST DATE FINDER
 * Scans Column A of 'Tracking' to find the exact row matching targetDate.
 * Compares Year, Month, and Day natively without relying on string formatting.
 */
function findTrackingRowForDate(sheet, targetDate) {
  const topTargetRow = 7;
  const lastRow = sheet.getLastRow();
  if (lastRow < topTargetRow) return -1;

  const t = new Date(targetDate);
  const targetY = t.getFullYear();
  const targetM = t.getMonth();
  const targetD = t.getDate();

  const values = sheet.getRange(topTargetRow, 1, (lastRow - topTargetRow) + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const cell = values[i][0];
    if (cell) {
      const d = new Date(cell);
      if (!isNaN(d.getTime())) {
        if (d.getFullYear() === targetY && d.getMonth() === targetM && d.getDate() === targetD) {
          return topTargetRow + i;
        }
      }
    }
  }
  return -1;
}
