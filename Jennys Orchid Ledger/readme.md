# 🌸 Jenny's Orchid Management System (OMS)

> A centralized digital ledger and automation suite for managing orchid collection records, care schedules, bloom cycles, and Google Drive photo archives.

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google-drive&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![License: Private](https://img.shields.io/badge/License-Private-red?style=for-the-badge)

---

## 📂 System File Architecture

The backend automation and frontend user interfaces are organized into a modular, collision-free Google Apps Script structure:

* **`Config.gs`** — Global constants, active workbook ID, and column index mappings.
* **`MenuAndNavigation.gs`** — Custom `🌸 Orchid Tools` toolbar menu, web app dashboard launcher, sidebar navigation bridge, and summary formulas.
* **`InventoryAndArchive.gs`** — Two-stage specimen provisioning (`Provision` & `Finalize`), decommissioning/tab recycling (`TMP_` protocol), and Drive photo folder sync.
* **`EventHandlers.gs`** — Central `onEdit` trigger handling Maintenance Log synchronization, Inventory bloom status updates, and automatic duration pairing.
* **`CareAndRepotTools.gs`** — Bulk watering dialog backend, historical date sync by YYYY-MM-DD, fertilizer note appending, log deduplication, and repot preview engine.
* **`FormattingAndDiagnostics.gs`** — Bloom status styling (`D5`), repot stoplights (`D10`), date standardization, Excel (`.xlsx`) export utility, and system health checks.
* **`WebAppBackend.gs`** — Container-bound web app HTTP `GET` endpoint and backend bridge for *The Potted Ledger* dashboard.
* **`Index.html`** — Dark-mode responsive frontend UI featuring collection filtering, specimen passports, photo galleries, and chronological log editors.
* **`Sidebar.html`** — Quick-jump navigation drawer for sheets and specimen names.
* **`HelpFile.html`** — In-app User Guide modal (`System Manual & Guide`).
* **`WateringDialog.html`** — Interactive modeless popup for multi-row bulk watering entries.
* **`repot-preview.html`** — Calculation audit modal for reviewing repotting schedules.
* **`ChangeLog.html`** — Version history reference.
* **`ReadMe.gs`** — System documentation archive.

---

## 📊 Data Flow & Cell Mapping (Compact A:N Layout)

* **Master Inventory (`Inventory`):**
  * **Col A:** ID (Hyperlinked to individual specimen tabs)
  * **Col B:** Orchid Name | **Col C:** Genus | **Col I:** Acquisition Date
  * **Col K:** Bloom Status | **Col L:** Last Bloom Date | **Col O:** Target Foot Candles
  * **Col T:** Repot Every | **Col U:** Next Repot Due

* **Individual Specimen Tabs (`1`, `2`, `3`, ...):**
  * **`B2:B6`** — Botanical profile (Genus, Lineage, Alliance, Name, Native Range)
  * **`B11:B16`** — Care baseline (Light, Watering, Humidity, Media, Fertilizer)
  * **`D5`** — Live Bloom Status badge (`🌸 In Bloom`, `🌱 In Spike`, `🌼 In Bud`, `⬜ Not in Bloom`)
  * **`D10`** — Next Repot Due date with color-coded stoplight alert (Green / Amber / Red)
  * **`A20:D`** — Bloom Log (In Bloom, Out of Bloom, Duration, Notes)
  * **`F20:G`** — Observation Notes (Date, Notes)
  * **`I20:J`** — Repotting Log (Date, Notes)
  * **`L20:N`** — Watering Log (Date, Method, Notes)

---

## 🚀 Quick Setup & Deployment

1. **Authorize Scripts:** Open your spreadsheet and run `🌸 Orchid Tools > 🏥 Run System Health Check` to grant one-time script permissions.
2. **Build Drive Folders:** Run `🌸 Orchid Tools > 🖼️ Photo Archive & Drive > 📁 Create / Sync All Drive Photo Folders` to generate your master `Orchid Photos` repository and plant subfolders.
3. **Launch Web App Dashboard:** 
   * Go to `Extensions > Apps Script > Deploy > New deployment > Web app` (Execute as: *Me*, Who has access: *Only myself*).
   * Copy your generated Web app URL (ending in `/exec`), and paste it into `🌸 Orchid Tools > 🔗 Set Web App URL`.
   * Open Jenny's dashboard anytime via `🌸 Orchid Tools > 🌐 Open Web App Dashboard`.

---
*Maintained securely for Jenny's Botanical Collection Repository.*
