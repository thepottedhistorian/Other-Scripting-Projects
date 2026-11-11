/**
 * ============================================================================
 * Google Apps Script: Drive Folder Indexer
 * ============================================================================
 * Description: Scans designated Google Drive folders and populates file metadata
 *              into specific target tabs ('Lab Work' and 'Discharge Summaries').
 * ============================================================================
 */

/**
 * Main execution function to update both document sheets using Script Properties.
 */
function updateAllDriveIndexes() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const labWorkFolderId = scriptProperties.getProperty('FOLDER_ID_LAB_WORK');
  const dischargeFolderId = scriptProperties.getProperty('FOLDER_ID_DISCHARGE_SUMMARIES');

  if (!labWorkFolderId || !dischargeFolderId) {
    throw new Error('Missing folder IDs in Script Properties. Ensure FOLDER_ID_LAB_WORK and FOLDER_ID_DISCHARGE_SUMMARIES are set.');
  }

  // 1. Index 'Lab Work' folder
  importFolderContents('Lab Work', labWorkFolderId);

  // 2. Index 'Discharge Summaries' folder
  importFolderContents('Discharge Summaries', dischargeFolderId);
}

/**
 * Helper function that processes a single tab and folder pair.
 * 
 * @param {string} sheetName - Exact name of the tab in your spreadsheet.
 * @param {string} folderId  - Alphanumeric ID of the target Google Drive folder.
 */
function importFolderContents(sheetName, folderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  // Create the tab if it doesn't exist yet
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();

    // Reset sheet content and write headers
    sheet.clear();
    sheet.appendRow(['File Name', 'File URL', 'Last Updated']);
    sheet.getRange('A1:C1').setFontWeight('bold');

    const data = [];
    while (files.hasNext()) {
      const file = files.next();
      data.push([
        file.getName(),
        file.getUrl(),
        file.getLastUpdated()
      ]);
    }

    // Write all row data in a single batch for performance
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, 3).setValues(data);
    }
  } catch (error) {
    Logger.log(`Error indexing folder [${folderId}] into sheet [${sheetName}]: ${error.toString()}`);
  }
}

/* 
  Make sure your existing importFountainData() and 
  importZoeWeightsFromEmail() functions remain in the script below! 
*/
