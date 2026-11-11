/**
 * =============================================================================
 * SCRIPT: Daily Midnight Rollover for "Tracking"
 * AUTHOR: Rebecca Short
 * SYSTEM: Zöe's Daily Care Tracker
 * 
 * EXACT PHYSICAL COLUMN MAPPING (Header Row 6):
 * A (1): Date
 * B (2): Sub-Q (ml)
 * C (3): Slurry & Hydracare (ml)
 * D (4): Fountain (ml)
 * E (5): Weight (lbs)
 * F (6): K+ gel
 * G (7): Fish Oil (ml)
 * H (8): Phosphate Binder
 * I (9): Kidney Support Gold [ACTIVE CHECKBOX]
 * J (10): Varenzin CA-1 [HISTORICAL - LEAVE BLANK]
 * K (11): Prednisone [HISTORICAL - LEAVE BLANK]
 * L (12): Darbepoetin [ACTIVE CHECKBOX]
 * M (13): Amlodipine
 * N (14): Total Intake (ml) [FORMULA: =SUM(B:D)]
 * O (15): Other Medications Given
 * P (16): Notes
 * =============================================================================
 */

function addMidnightRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var today = new Date();
  var currentDateString = Utilities.formatDate(today, timeZone, "yyyy-MM-dd");

  var trackSheet = ss.getSheetByName("Tracking");
  if (!trackSheet) {
    Logger.log("Error: 'Tracking' sheet could not be found.");
    return;
  }

  // Row 7 is the top data row directly under Row 6 headers
  var targetRow = 7;
  var scanDepth = 5;

  // 1. DUPLICATE DATE DETECTION
  var topDates = trackSheet.getRange(targetRow, 1, scanDepth, 1).getValues();
  var todayRowIndex = -1;

  for (var i = 0; i < topDates.length; i++) {
    if (topDates[i][0]) {
      var entryDate = new Date(topDates[i][0]);
      if (!isNaN(entryDate.getTime())) {
        var entryDateString = Utilities.formatDate(entryDate, timeZone, "yyyy-MM-dd");
        if (entryDateString === currentDateString) {
          todayRowIndex = targetRow + i;
          break;
        }
      }
    }
  }

  // 2. INSERT ROW IF NOT FOUND
  if (todayRowIndex === -1) {
    trackSheet.insertRowBefore(targetRow);
    
    // Set Date in Col A
    var newDateCell = trackSheet.getRange(targetRow, 1);
    newDateCell.setValue(today);
    newDateCell.setNumberFormat("dddd, mmmm d, yyyy");
    
    // Checkboxes: 
    // Col I (9)  = Kidney Support Gold
    // Col L (12) = Darbepoetin
    var checkboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    trackSheet.getRange(targetRow, 9).setDataValidation(checkboxValidation);  // Col I (Kidney Support Gold)
    trackSheet.getRange(targetRow, 12).setDataValidation(checkboxValidation); // Col L (Darbepoetin)
    
    // Ensure hidden historical columns J & K do NOT get checkboxes
    trackSheet.getRange(targetRow, 10).clearDataValidations();
    trackSheet.getRange(targetRow, 11).clearDataValidations();
    
    // Total Intake Formula in Col N (14): Sub-Q (B) + Slurry (C) + Fountain (D)
    trackSheet.getRange(targetRow, 14).setFormula(
      '=SUM(B' + targetRow + ':D' + targetRow + ')'
    );
    
    Logger.log("Inserted new tracking row for " + currentDateString + " at Row " + targetRow);
  } else {
    // Reconcile Checkboxes on active row
    var checkboxValidation = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    trackSheet.getRange(todayRowIndex, 9).setDataValidation(checkboxValidation);
    trackSheet.getRange(todayRowIndex, 12).setDataValidation(checkboxValidation);

    // Reconcile Formula in Col N (14): Sub-Q (B) + Slurry (C) + Fountain (D)
    var formulaCell = trackSheet.getRange(todayRowIndex, 14);
    formulaCell.setFormula(
      '=SUM(B' + todayRowIndex + ':D' + todayRowIndex + ')'
    );
    Logger.log("Row and formula verified at Row " + todayRowIndex);
  }
}
