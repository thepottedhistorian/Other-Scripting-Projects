/**
 * =============================================================================
 * JENNY'S ORCHID MANAGEMENT SYSTEM - TECHNICAL README (ReadMe.gs)
 * =============================================================================
 * PROJECT: Jenny's Orchid Collection Ledger
 * VERSION: 1.0.0 (Compact A:N Edition)
 * UPDATED: 2026-11-13
 * ARCHITECTURE: Google Apps Script (Container-Bound Workbook Engine)
 * =============================================================================
 *
 * 1. SYSTEM PURPOSE
 * -----------------------------------------------------------------------------
 * A centralized digital ledger for managing orchid collection records, care
 * schedules, bloom cycles, and historical specimen archives. Designed for
 * seamless day-to-day tracking directly inside Google Sheets as well as
 * companion web app access.
 *
 * CORE CAPABILITIES:
 * - Two-Stage Specimen Onboarding: Auto-assigns sequential ID numbers and
 *   generates formatted specimen ledger sheets from the Template tab.
 * - Automated Care & Watering Sync: Transcribes entries from the central
 *   Maintenance Log into individual specimen sheets (Columns L:N).
 * - Dynamic Repot Scheduling: Calculates Next Repot Due dates using the
 *   upper bound of each orchid's repotting frequency (e.g., "1-2 years" = 24 mos).
 * - Bloom Cycle Tracking: Auto-logs bloom start/end dates, calculates exact
 *   bloom durations, and applies color-coded status badges to Cell D5.
 * - Sheet Recycling & Archiving: Moves retired orchids to Archived Orchids,
 *   computes total collection age, and recycles tabs as TMP_[ID] for reuse.
 *
 * 2. CONSOLIDATED FILE DIRECTORY
 * -----------------------------------------------------------------------------
 * - Config.gs: Global constants, active workbook ID, and column index mappings.
 * - MenuAndNavigation.gs: Custom "🌸 Orchid Tools" toolbar menu, sidebar
 *   navigation bridge, and collection count/watering summary functions.
 * - InventoryAndArchive.gs: Specimen provisioning (Stage 1 & Stage 2),
 *   decommissioning/recycling engine, and repot due date push utilities.
 * - EventHandlers.gs: Central onEdit trigger handling Maintenance Log sync,
 *   Inventory bloom status updates, and automatic duration recalculation.
 * - CareAndRepotTools.gs: Bulk watering dialog backend, historical date sync,
 *   fertilizer note appending, log deduplication, and repot preview engine.
 * - FormattingAndDiagnostics.gs: Bloom status styling (D5), repot stoplights
 *   (D10), date standardization, Excel (.xlsx) export, and health checks.
 * - Sidebar.html: Quick-jump navigation drawer for sheets and specimen names.
 * - HelpFile.html: In-app User Guide modal ("System Manual & Guide").
 * - WateringDialog.html: Interactive modeless popup for bulk watering entries.
 * - repot-preview.html: Calculation audit modal before applying repot dates.
 * - ChangeLog.html: Version history reference.
 * - ReadMe.gs: System architecture and operational documentation.
 *
 * 3. WORKBOOK DATA FLOW & CELL MAPPING (COMPACT A:N LAYOUT)
 * -----------------------------------------------------------------------------
 * - Master Inventory ("Inventory"):
 *   Col A: ID (Hyperlinked to ID tab) | Col B: Orchid Name | Col C: Genus
 *   Col I: Acquisition Date | Col K: Bloom Status | Col L: Last Bloom Date
 *   Col O: Target Foot Candles | Col T: Repot Every | Col U: Next Repot Due
 *
 * - Individual Specimen Tabs ("1", "2", ...):
 *   B2:B6   -> Botanical profile (Genus, Lineage, Alliance, Name, Native Range)
 *   B11:B16 -> Care baseline (Light, Watering, Humidity, Media, Fertilizer)
 *   D5      -> Live Bloom Status badge (🌸 In Bloom, 🌱 In Spike, 🌼 In Bud, ⬜ Not in Bloom)
 *   D10     -> Next Repot Due date with stoplight alert (Green / Amber / Red)
 *   A20:D   -> Bloom Log (In Bloom, Out of Bloom, Duration, Notes)
 *   F20:G   -> Observation Notes (Date, Notes)
 *   I20:J   -> Repotting Log (Date, Notes)
 *   L20:N   -> Watering Log (Date, Method, Notes)
 *
 * 4. MAINTENANCE & BEST PRACTICES
 * -----------------------------------------------------------------------------
 * - Never delete numbered orchid tabs manually; always use
 *   "🌸 Orchid Tools > 📦 Archive Orchid" so the tab is cleanly recycled as TMP_[ID].
 * - Run "🌸 Orchid Tools > 🏥 Run System Health Check" periodically to verify
 *   that all active Inventory rows have matching specimen tabs and valid dates.
 * =============================================================================
 */
