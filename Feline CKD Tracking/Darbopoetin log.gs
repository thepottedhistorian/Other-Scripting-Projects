/**
 * =========================================================================
 * SCRIPT: Zöe's Master Care Log & Darbepoetin 21-Day Router
 * AUTHOR: Rebecca Short
 * VERSION: 10.1 (Strict 21-Day Cycle Rollover & Status Update Fix)
 * =========================================================================
 */

/**
 * WATCHER: Listens for Column L (Column 12) checks on "Tracking" sheet.
 * Stamps current time into "Darbepoetin Log", updates status, and increments dose count.
 */
function onEdit(e) {
  if (!e) return;

  const range = e.range;
  const sheet = range.getSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Core Filter: Must be "Tracking" sheet, Column L (12), and checked 'true'
  if (sheet.getName() !== "Tracking" || range.getColumn() !== 12 || range.getValue() !== true) {
    return;
  }

  // Prevent double execution on rapid clicks
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
    // 1. Stamp time in Column B
    const timeCell = logSheet.getRange(targetRow, 2);
    timeCell.setNumberFormat("@");
    timeCell.setValue(timeStr);
    timeCell.setFontWeight("normal");
    timeCell.setHorizontalAlignment("left");

    // 2. Update Phase in Column C to 'Dose Given Today'
    const phaseCell = logSheet.getRange(targetRow, 3);
    phaseCell.setValue("Dose Given Today");
    phaseCell.setFontStyle("italic");

    // 3. Increment dose count in Column E from previous day
    const prevDoses = Number(logSheet.getRange(targetRow + 1, 5).getValue()) || 0;
    logSheet.getRange(targetRow, 5).setValue(prevDoses + 1);

    ss.toast("Darbepoetin logged at " + timeStr, "Success");
  } else {
    ss.toast("Error: Could not find '" + targetDateStr + "' on Darbepoetin Log.", "Sync Error");
  }

  lock.releaseLock();
}

/**
 * AUTOMATION: Daily Darbepoetin Midnight Rollover (Strict 21-Day Cycle)
 * Adds today's row, increments cycle day (1 to 21), flags missed doses,
 * and alerts when Day 1 (Dose Due) arrives.
 */
function dailyDarbepoetinRollover() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Darbepoetin Log") || ss.getSheetByName("Darbopoetin Log");
  if (!logSheet) return;

  var today = new Date();
  var dateString = Utilities.formatDate(today, ss.getSpreadsheetTimeZone(), "EEEE, MMMM d, yyyy");

  var topDate = logSheet.getRange(2, 1).getValue();
  var topDateFormatted = Utilities.formatDate(new Date(topDate), ss.getSpreadsheetTimeZone(), "EEEE, MMMM d, yyyy");

  if (topDateFormatted !== dateString) {
    var yesterdayCycleDay = Number(logSheet.getRange(2, 4).getValue());
    var yesterdayTime = logSheet.getRange(2, 2).getValue();

    // If yesterday was Day 1 (Dose Due) and no time was stamped, mark MISSED DOSE
    if (yesterdayCycleDay === 1 && (yesterdayTime === "" || yesterdayTime === null)) {
      var missedCell = logSheet.getRange(2, 2);
      missedCell.setValue("MISSED DOSE");
      missedCell.setFontWeight("bold");
      missedCell.setHorizontalAlignment("center");
    }

    // Insert new top row at Row 2
    logSheet.insertRowBefore(2);
    logSheet.getRange(2, 1).setValue(dateString);

    var prevCycleDay = Number(logSheet.getRange(3, 4).getValue()) || 0;
    var prevDoses = Number(logSheet.getRange(3, 5).getValue()) || 0;

    // Strict 21-Day Logic: counts Day 1 to 21, then rolls over to Day 1 (Dose Due)
    var calculatedCycleDay = (prevCycleDay >= 21) ? 1 : prevCycleDay + 1;
    var calculatedPhase = (calculatedCycleDay === 1) ? "Dose Due Today" : "21-Day Cycle";

    var phaseCell = logSheet.getRange(2, 3);
    phaseCell.setValue(calculatedPhase);
    if (calculatedPhase === "Dose Due Today") {
      phaseCell.setFontStyle("italic");
    } else {
      phaseCell.setFontStyle("normal");
    }

    logSheet.getRange(2, 4).setValue(calculatedCycleDay);
    logSheet.getRange(2, 5).setValue(prevDoses);

    // Send alert on Dose Day (Day 1)
    if (calculatedCycleDay === 1) {
      var recipient = Session.getActiveUser().getEmail();
      var subject = "💉 Darbepoetin Alert: Zöe's Dose Due Today!";
      var body = "Hi Rebecca,\n\nToday (" + dateString + ") is Day 1 of Zöe's 21-day Darbepoetin cycle. Her injection is due today.\n\n" +
                 "Please remember to administer her 0.2 mL dose and check off the box on the Tracking sheet when completed.\n\n" +
                 "— Zöe's Care Log Assistant";

      MailApp.sendEmail(recipient, subject, body);
    }
  }
}

/**
 * REPAIR UTILITY: Rebuilds rows 2 through 12 to match the Oct 28 dose anchor.
 * Sets Today (Saturday, Nov 7) to Day 11 of 21. Next dose: Wednesday, Nov 18 (Day 1).
 */
function fixCurrentDarbepoetinLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Darbepoetin Log") || ss.getSheetByName("Darbopoetin Log");
  if (!logSheet) return;

  const rowsData = [
    { row: 2,  dateStr: "Saturday, November 7, 2026",  time: "",         phase: "21-Day Cycle",     day: 11, doses: 6 },
    { row: 3,  dateStr: "Friday, November 6, 2026",    time: "",         phase: "21-Day Cycle",     day: 10, doses: 6 },
    { row: 4,  dateStr: "Thursday, November 5, 2026",  time: "",         phase: "21-Day Cycle",     day: 9,  doses: 6 },
    { row: 5,  dateStr: "Wednesday, November 4, 2026", time: "",         phase: "21-Day Cycle",     day: 8,  doses: 6 },
    { row: 6,  dateStr: "Tuesday, November 3, 2026",   time: "",         phase: "21-Day Cycle",     day: 7,  doses: 6 },
    { row: 7,  dateStr: "Monday, November 2, 2026",    time: "",         phase: "21-Day Cycle",     day: 6,  doses: 6 },
    { row: 8,  dateStr: "Sunday, November 1, 2026",    time: "",         phase: "21-Day Cycle",     day: 5,  doses: 6 },
    { row: 9,  dateStr: "Saturday, October 31, 2026",  time: "",         phase: "21-Day Cycle",     day: 4,  doses: 6 },
    { row: 10, dateStr: "Friday, October 30, 2026",    time: "",         phase: "21-Day Cycle",     day: 3,  doses: 6 },
    { row: 11, dateStr: "Thursday, October 29, 2026",  time: "",         phase: "21-Day Cycle",     day: 2,  doses: 6 },
    { row: 12, dateStr: "Wednesday, October 28, 2026", time: "12:44 PM", phase: "Dose Given Today", day: 1,  doses: 6 }
  ];

  rowsData.forEach(item => {
    logSheet.getRange(item.row, 1).setValue(item.dateStr);

    const timeCell = logSheet.getRange(item.row, 2);
    timeCell.setValue(item.time);
    if (item.time === "MISSED DOSE") {
      timeCell.setFontWeight("bold").setHorizontalAlignment("center");
    } else {
      timeCell.setFontWeight("normal").setHorizontalAlignment("left");
    }

    const phaseCell = logSheet.getRange(item.row, 3);
    phaseCell.setValue(item.phase);
    if (item.phase === "Dose Due Today" || item.phase === "Dose Given Today") {
      phaseCell.setFontStyle("italic");
    } else {
      phaseCell.setFontStyle("normal");
    }

    logSheet.getRange(item.row, 4).setValue(item.day);
    logSheet.getRange(item.row, 5).setValue(item.doses);
  });

  ss.toast("Darbepoetin log realigned: Today is Day 11 of 21. Next dose: Nov 18.", "Success");
}
