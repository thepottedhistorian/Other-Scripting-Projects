/**
 * =============================================================================
 * JENNY'S ORCHID LEDGER - GLOBAL CONFIGURATION (Config.gs)
 * =============================================================================
 * Version: 1.1.0 (Tailored Edition + Drive Photo Folders)
 * Purpose: Centralized spreadsheet constants, column mappings, and Drive settings.
 * =============================================================================
 */

const MASTER_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

const CONFIG = {
  DATABASE_ID: MASTER_ID,
  ARCHIVE_PREFIX: "TMP_",
  PHOTOS_FOLDER_NAME: "Orchid Photos",
  VERSION: "1.1.0",
  PROJECT_NAME: "Jenny's Orchid Management System"
};

/**
 * 1. INVENTORY COLUMN INDICES (0-indexed)
 */
const ID_COL = 0;           // Column A: ID
const NAME_COL = 1;          // Column B: Orchid Name
const GENUS_COL = 2;         // Column C: Genus
const ACQ_COL = 8;           // Column I: Acquisition Date
const REPOT_NOTES_COL = 9;   // Column J: Source / Notes
const BLOOM_STATUS_COL = 10; // Column K: In Bloom Status
const LAST_BLOOM_COL = 11;   // Column L: Last Bloom Date
const LIGHT_FC_COL = 14;     // Column O: Target Foot Candles
const FREQ_COL = 19;         // Column T: Repot Every
const NEXT_COL = 20;         // Column U: Next Repot Due

/**
 * 2. CORE WORKBOOK SHEET NAMES
 */
const INVENTORY_SHEET = "Inventory";
const ARCHIVE_SHEET = "Archived Orchids";
const MAINT_LOG_SHEET = "Maintenance Log";
const TEMPLATE_SHEET = "Template";
