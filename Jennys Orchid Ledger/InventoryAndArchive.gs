/**
 * =============================================================================
 * JENNY'S ORCHID LEDGER - PROVISIONING & ARCHIVE (InventoryAndArchive.gs)
 * =============================================================================
 * Purpose: Manages new specimen onboarding (2-stage provisioning), specimen
 *          archiving with TMP_ sheet recycling, and repot due date delivery.
 * =============================================================================
 */

function provisionNewOrchid() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const ui = SpreadsheetApp.getUi();

  if (!inv) {
    ui.alert("Error: 'Inventory' sheet not found.");
    return;
  }

  const lastRow = inv.getLastRow();
  let targetRow = -1;

  for (let i = 2; i <= lastRow; i++) {
    if (inv.isRowHiddenByUser(i)) {
      targetRow = i;
      break;
    }
  }

  const data = inv.getDataRange().getValues();
  const activeIDs = new Set();
  for (let i = 1; i < data.length; i++) {
    if (inv.isRowHiddenByUser(i + 1)) continue;
    const idVal = data[i][0];
    if (idVal !== "" && idVal !== null && !isNaN(idVal)) {
      activeIDs.add(parseInt(idVal, 10));
    }
  }

  let nextAvailableID = 1;
  while (activeIDs.has(nextAvailableID)) {
    nextAvailableID++;
  }

  if (targetRow !== -1) {
    inv.showRows(targetRow);
    inv.getRange(targetRow, 1).setValue(nextAvailableID);
    inv.getRange(targetRow, 2).activate();
    ui.alert(`Recycled row ${targetRow} restored and auto-assigned ID ${nextAvailableID}.\n\nFill out the botanical details and then run '2. Finalize & Link Ledger'.`);
  } else {
    targetRow = lastRow + 1;
    inv.getRange(targetRow, 1).setValue(nextAvailableID);
    inv.getRange(targetRow, 2).activate();
    ui.alert(`Ready on row ${targetRow} with auto-assigned ID ${nextAvailableID}.\n\nFill out the botanical details and then run '2. Finalize & Link Ledger'.`);
  }
}

function finalizeNewOrchid() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const template = ss.getSheetByName(TEMPLATE_SHEET) || ss.getSheetByName("template");
  const maintenanceLog = ss.getSheetByName(MAINT_LOG_SHEET);
  const ui = SpreadsheetApp.getUi();

  const activeRow = inv.getActiveCell().getRow();
  if (activeRow < 2) {
    ui.alert("Please select a specimen row in Inventory first.");
    return;
  }

  const maxCol = Math.min(inv.getLastColumn(), 22);
  const data = inv.getRange(activeRow, 1, 1, maxCol).getValues()[0];
  const orchidId = data[0];
  const orchidName = data[1];

  if (!orchidId || orchidId === "") {
    ui.alert("Critical Error: No ID found in Column A.");
    return;
  }

  if (ss.getSheetByName(String(orchidId))) {
    ui.alert(`Error: A sheet for ID ${orchidId} already exists.`);
    return;
  }

  if (!template) {
    ui.alert("Error: 'Template' sheet not found.");
    return;
  }

  // Reuse a recycled TMP_ tab if available, otherwise copy Template
  const tmpSheet = ss.getSheets().find(s => s.getName().toUpperCase().startsWith("TMP_"));
  let newSheet;
  if (tmpSheet) {
    newSheet = tmpSheet.setName(String(orchidId));
  } else {
    newSheet = template.copyTo(ss).setName(String(orchidId));
  }
  newSheet.showSheet();

  // Calculate initial repot due date
  const acqRaw = data[ACQ_COL];
  const freqRaw = data[FREQ_COL];
  let acqDate = (acqRaw instanceof Date && !isNaN(acqRaw.getTime())) ? acqRaw : (acqRaw ? new Date(acqRaw) : null);
  if (acqDate && isNaN(acqDate.getTime())) acqDate = null;

  const freqMonths = (typeof parseFrequencyToMonths === 'function' ? parseFrequencyToMonths(freqRaw) : null) || 12;
  let nextRepotDate = null;

  if (acqDate) {
    nextRepotDate = new Date(acqDate);
    nextRepotDate.setMonth(nextRepotDate.getMonth() + freqMonths);
    inv.getRange(activeRow, NEXT_COL + 1).setValue(nextRepotDate).setNumberFormat('MMMM d, yyyy');
  } else if (data[NEXT_COL]) {
    const parsed = new Date(data[NEXT_COL]);
    if (!isNaN(parsed.getTime())) nextRepotDate = parsed;
  }

  // Populate compact A:N header cards (formulas already point to Inventory, fallback values for static templates)
  if (!newSheet.getRange("B2").getFormula()) newSheet.getRange("B2").setValue(data[2] || "");
  if (!newSheet.getRange("B3").getFormula()) newSheet.getRange("B3").setValue(data[5] || "");
  if (!newSheet.getRange("B4").getFormula()) newSheet.getRange("B4").setValue(data[6] || "");
  if (!newSheet.getRange("B5").getFormula()) newSheet.getRange("B5").setValue(data[1] || "");
  if (!newSheet.getRange("B6").getFormula()) newSheet.getRange("B6").setValue(data[7] || "");

  newSheet.getRange("D5").setValue(data[10] || "Not in Bloom");
  if (typeof applyBloomStatusFormatting === 'function') {
    applyBloomStatusFormatting(newSheet);
  }

  if (nextRepotDate) {
    const d10 = newSheet.getRange("D10");
    d10.setValue(nextRepotDate).setNumberFormat('MMMM d, yyyy').setHorizontalAlignment("center");
  }

  const sheetId = newSheet.getSheetId();
  inv.getRange(activeRow, 1).setFormula(`=HYPERLINK("#gid=${sheetId}", "${orchidId}")`);

  if (maintenanceLog) {
    const nextLogRow = Math.max(maintenanceLog.getLastRow() + 1, 3);
    maintenanceLog.getRange(nextLogRow, 1).setValue(orchidId);
    maintenanceLog.getRange(nextLogRow, 2).setValue(orchidName).setFontStyle("italic");
    if (nextRepotDate) {
      maintenanceLog.getRange(nextLogRow, 7).setValue(nextRepotDate).setNumberFormat('MMMM d, yyyy');
    }
  }

  ui.alert(`Success! Specimen sheet and Maintenance Log entry created for ${orchidName} (ID: ${orchidId}).`);
}

function archiveOrchidById_(orchidId, batchStatus = null, batchNotes = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}

  const inventory = ss.getSheetByName(INVENTORY_SHEET);
  if (!inventory) return false;

  let archiveSheet = ss.getSheetByName(ARCHIVE_SHEET);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(ARCHIVE_SHEET);
    archiveSheet.getRange("A1:J1").setValues([[
      "Orchid ID", "Name", "Genus", "Species / Hybrid", "Acquisition Date",
      "Archived Date", "Final Status", "Age", "Last Repot Date", "Notes"
    ]]);
    archiveSheet.setFrozenRows(1);
  }

  const data = inventory.getDataRange().getValues();
  let rowIndex = -1;
  let rowValues = [];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][ID_COL]).trim() === String(orchidId).trim()) {
      rowIndex = i + 1;
      rowValues = data[i];
      break;
    }
  }

  if (rowIndex === -1) {
    if (!batchStatus && ui) ui.alert("Orchid ID " + orchidId + " not found in Inventory.");
    return false;
  }

  let finalStatus = batchStatus;
  let notes = batchNotes;

  if (batchStatus === null && ui) {
    const statusResp = ui.prompt("Archive Orchid " + orchidId, "Enter final status (e.g., Died, Gifted, Rehomed):", ui.ButtonSet.OK_CANCEL);
    if (statusResp.getSelectedButton() !== ui.Button.OK) return false;
    finalStatus = statusResp.getResponseText().trim() || "Archived";

    const notesResp = ui.prompt("Archive Orchid " + orchidId, "Optional notes:", ui.ButtonSet.OK_CANCEL);
    if (notesResp.getSelectedButton() !== ui.Button.OK) return false;
    notes = notesResp.getResponseText().trim();
  }

  const orchidSheet = ss.getSheetByName(String(orchidId).trim());
  let lastRepotDate = null;
  if (orchidSheet && orchidSheet.getLastRow() >= 20) {
    lastRepotDate = getMostRecentDateFromRange_(orchidSheet.getRange(20, 9, orchidSheet.getLastRow() - 19, 1));
  }

  const today = new Date();
  const acqRaw = rowValues[ACQ_COL];
  const acqDate = (acqRaw instanceof Date && !isNaN(acqRaw.getTime())) ? acqRaw : (acqRaw ? new Date(acqRaw) : null);
  const ageStr = (acqDate && !isNaN(acqDate.getTime()) && typeof calculateDurationString_ === 'function')
    ? calculateDurationString_(acqDate, today)
    : "";

  archiveSheet.appendRow([
    rowValues[ID_COL],
    rowValues[NAME_COL],
    rowValues[GENUS_COL],
    rowValues[5] || "",
    rowValues[ACQ_COL] || "",
    Utilities.formatDate(today, tz, "M/d/yyyy"),
    finalStatus,
    ageStr,
    lastRepotDate ? Utilities.formatDate(lastRepotDate, tz, "M/d/yyyy") : "N/A",
    notes || ""
  ]);

  inventory.getRange(rowIndex, 1, 1, inventory.getLastColumn()).clearContent();
  inventory.hideRows(rowIndex);

  if (orchidSheet) {
    const templateSheet = ss.getSheetByName(TEMPLATE_SHEET) || ss.getSheetByName("template");
    if (templateSheet) {
      orchidSheet.clear();
      templateSheet.getDataRange().copyTo(orchidSheet.getRange(1, 1));
      try {
        orchidSheet.setName("TMP_" + orchidId);
      } catch (e) {
        orchidSheet.setName("RECYCLED_" + new Date().getTime());
      }
      orchidSheet.hideSheet();
    }
  }

  return true;
}

function getMostRecentDateFromRange_(range) {
  const values = range.getValues();
  let latest = null;
  for (let r = 0; r < values.length; r++) {
    const v = values[r][0];
    if (v instanceof Date && !isNaN(v.getTime())) {
      if (!latest || v.getTime() > latest.getTime()) latest = v;
    }
  }
  return latest;
}

function archiveOrchid() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Archive Orchid", "Enter the Orchid ID to archive:", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const orchidId = response.getResponseText().trim();
  if (orchidId && archiveOrchidById_(orchidId)) {
    ui.alert(`ID ${orchidId} has been moved to Archived Orchids and its tab is ready for reuse.`);
  }
}

function bulkArchiveOrchids() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Bulk Archive", "Enter Orchid IDs separated by commas:", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const ids = response.getResponseText().split(",").map(id => id.trim()).filter(id => id.length > 0);
  if (ids.length === 0) return;

  const batchStatus = ui.prompt("Batch Archive", "Final status for all (e.g., Gifted):", ui.ButtonSet.OK).getResponseText() || "Archived";
  const batchNotes = ui.prompt("Batch Archive", "Notes for all:", ui.ButtonSet.OK).getResponseText() || "";

  let count = 0;
  ids.forEach(id => {
    if (archiveOrchidById_(id, batchStatus, batchNotes)) count++;
  });

  ui.alert(`Successfully archived ${count} orchids.`);
}

function pushInventoryTToOrchidD10() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  if (!inv) return;

  const invData = inv.getDataRange().getValues();
  let updateCount = 0;

  for (let i = 1; i < invData.length; i++) {
    const id = invData[i][ID_COL];
    const nextDue = invData[i][NEXT_COL];

    if (id && nextDue instanceof Date && !isNaN(nextDue.getTime())) {
      const orchidSheet = ss.getSheetByName(String(id).trim());
      if (orchidSheet) {
        const cellD10 = orchidSheet.getRange("D10");
        cellD10.clearContent();
        cellD10.setValue(nextDue).setNumberFormat('MMMM d, yyyy').setHorizontalAlignment("center").setVerticalAlignment("middle");
        updateCount++;
      }
    }
  }

  try {
    SpreadsheetApp.getUi().alert(`Success: ${updateCount} orchid sheets updated. D10 now matches Inventory.`);
  } catch (e) {}
}

/**
 * =============================================================================
 * GOOGLE DRIVE PHOTO FOLDER GENERATOR & CARD SYNC
 * =============================================================================
 * Creates a master "Orchid Photos" folder in Jenny's Google Drive (if missing),
 * generates "[ID] - [Orchid Name]" subfolders for every active orchid, and
 * updates the Specimen Photo Archive card (J1:N5) on each numbered ID sheet.
 * =============================================================================
 */

function getOrCreateMasterPhotosFolder_() {
  const folderName = (typeof CONFIG !== 'undefined' && CONFIG.PHOTOS_FOLDER_NAME)
    ? CONFIG.PHOTOS_FOLDER_NAME
    : "Orchid Photos";

  const existing = DriveApp.getFoldersByName(folderName);
  if (existing.hasNext()) {
    return existing.next();
  }
  return DriveApp.createFolder(folderName);
}

function sanitizeDriveFolderName_(name) {
  return String(name || "").replace(/[/\\:*?"<>|]/g, "").trim();
}

function ensureSpecimenDrivePhotoFolder_(id, orchidName, orchidSheet) {
  const parentFolder = getOrCreateMasterPhotosFolder_();
  const idStr = String(id).trim();
  const cleanName = sanitizeDriveFolderName_(orchidName);
  const desiredName = cleanName ? `${idStr} - ${cleanName}` : idStr;

  let matchedFolder = null;
  const subFolders = parentFolder.getFolders();
  while (subFolders.hasNext()) {
    const folder = subFolders.next();
    const fName = folder.getName().trim();
    if (fName === idStr || fName.startsWith(idStr + " - ") || fName.startsWith(idStr + "-")) {
      matchedFolder = folder;
      if (fName !== desiredName) {
        folder.setName(desiredName);
      }
      break;
    }
  }

  let createdNew = false;
  if (!matchedFolder) {
    matchedFolder = parentFolder.createFolder(desiredName);
    createdNew = true;
  }

  // Count photos & find latest upload
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  let photoCount = 0;
  let newestTime = 0;
  let latestInfo = "No photos uploaded yet";

  const files = matchedFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType().startsWith("image/")) {
      photoCount++;
      const updated = file.getLastUpdated().getTime();
      if (updated > newestTime) {
        newestTime = updated;
        const dateStr = Utilities.formatDate(file.getLastUpdated(), tz, "MMM d, yyyy");
        latestInfo = `${file.getName()} (${dateStr})`;
      }
    }
  }

  // Update the Specimen Photo Archive card (K2:N5) on the ID sheet if present
  if (orchidSheet) {
    const cardHeader = orchidSheet.getRange("J1").getDisplayValue().trim();
    if (cardHeader.includes("Specimen Photo Archive")) {
      const folderUrl = matchedFolder.getUrl();
      orchidSheet.getRange("K2:N2").setValue(matchedFolder.getName());
      orchidSheet.getRange("K3:N3")
        .setFormula(`=HYPERLINK("${folderUrl}", "📁 Open Drive Photo Folder ↗")`)
        .setFontColor("#2563eb")
        .setFontWeight("bold");
      orchidSheet.getRange("K4:N4").setValue(`${photoCount} ${photoCount === 1 ? "photo" : "photos"}`);
      orchidSheet.getRange("K5:N5").setValue(latestInfo);
    }
  }

  return { folder: matchedFolder, createdNew: createdNew, photoCount: photoCount };
}

/**
 * Menu Item 1: Creates missing Drive folders for all active orchids in Inventory
 * and updates the photo count & hyperlink cards (K2:N5) across all ID sheets.
 */
function createAndSyncAllDrivePhotoFolders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inv = ss.getSheetByName(INVENTORY_SHEET);
  const ui = SpreadsheetApp.getUi();

  if (!inv) {
    ui.alert("Error: 'Inventory' sheet not found.");
    return;
  }

  const parentFolder = getOrCreateMasterPhotosFolder_();
  const data = inv.getDataRange().getValues();
  let createdCount = 0;
  let syncedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const rowNum = i + 1;
    if (inv.isRowHiddenByUser(rowNum)) continue;

    const id = data[i][ID_COL];
    const name = data[i][NAME_COL];
    if (!id || String(id).trim() === "") continue;

    const orchidSheet = ss.getSheetByName(String(id).trim());
    const result = ensureSpecimenDrivePhotoFolder_(id, name, orchidSheet);
    if (result.createdNew) createdCount++;
    syncedCount++;
  }

  SpreadsheetApp.flush();
  ui.alert(
    "Google Drive Photo Sync Complete!",
    `• Master Folder: "${parentFolder.getName()}"\n` +
    `• New Specimen Folders Created: ${createdCount}\n` +
    `• Total Specimen Tabs Synced: ${syncedCount}\n\n` +
    `Tip: Click the "Open Drive Photo Folder ↗" link in cell K3 on any orchid's tab to upload photos directly to that plant's folder.`,
    ui.ButtonSet.OK
  );
}

/**
 * Menu Item 2: Opens the master "Orchid Photos" Drive folder in a new browser tab.
 */
function openMainPhotosFolderModal() {
  const parentFolder = getOrCreateMasterPhotosFolder_();
  const url = parentFolder.getUrl();

  const html = HtmlService.createHtmlOutput(
    `<html><body style="font-family: sans-serif; padding: 14px;">
      <p>Opening <strong>${parentFolder.getName()}</strong> in Google Drive...</p>
      <p><a href="${url}" target="_blank" onclick="google.script.host.close()">Click here if it does not open automatically ↗</a></p>
      <script>
        var win = window.open("${url}", "_blank");
        if (win) { google.script.host.close(); }
      </script>
    </body></html>`
  ).setWidth(360).setHeight(130);

  SpreadsheetApp.getUi().showModalDialog(html, "Google Drive Photo Repository");
}
