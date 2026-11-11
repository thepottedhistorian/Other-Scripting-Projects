# Feline CKD Care Tracker & Automation Suite

![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=googlesheets&logoColor=white)
![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![Google Drive](https://img.shields.io/badge/Google%20Drive-FFBA00?style=for-the-badge&logo=googledrive&logoColor=white)
![Hardware](https://img.shields.io/badge/Hardware-PetLibro%20RFID-008080?style=for-the-badge)
![Scale Telemetry](https://img.shields.io/badge/Scale-Litter--Robot%204-5C2D91?style=for-the-badge)
![JavaScript](https://img.shields.io/badge/JavaScript-ES6+-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Clinical Protocol](https://img.shields.io/badge/Clinical%20Protocol-Stage%20IV%20CKD-D9381E?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Active%20Palliative%20Care-success?style=for-the-badge)

A centralized clinical telemetry, palliative care tracking, and medication management suite built for Zöe Short (Stage IV Chronic Kidney Disease, Systemic Hypertension, Hypercalcemia, Nephrolithiasis, and Grade II/VI Heart Murmur).

The platform bridges automated IoT drinking measurements, smart litter scale weigh-ins, daily therapeutic dosing records, and longitudinal veterinary lab archives into a unified Google Sheets database and Apps Script web app dashboard.

---

## Architecture Overview

```
    +--------------------------------+       +---------------------------------+
    |   PetLibro RFID Dockstream     |       |         Litter-Robot 4          |
    |  Smart Water Fountain (mL/oz)  |       |    Weight Telemetry Sensors     |
    +---------------+----------------+       +----------------+----------------+
                    |                                         |
                    | Automated Ingestion                     | CSV Export / Ingestion
                    v                                         v
    +--------------------------------------------------------------------------+
    |                         Google Sheets Database                           |
    |  ----------------------------------------------------------------------  |
    |  * Tracking (Master Daily Log)         * Darbopoetin_Log (21-Day Cycle)  |
    |  * Water_Log (Fountain Bouts)          * Prednisone_Log & Antibiotic_Log |
    |  * Weight_Log (Scale Trajectory)       * Lab_Work & Discharge_Summaries  |
    |  * Reference_Sheet (Doses & Clinical Protocol Key)                       |
    +------------------------------------+-------------------------------------+
                                         |
                                         v
    +------------------------------------+-------------------------------------+
    |                   Google Apps Script Engine                              |
    |  ----------------------------------------------------------------------  |
    |  * Midnight row initialization & baseline rolling averages               |
    |  * Fluid intake summation (Sub-Q + Slurry + Fountain; 200+ mL goal)      |
    |  * Acute weight drop detection & Darbepoetin cycle status tracking       |
    |  * Google Drive folder indexing for veterinary lab reports               |
    |  * Standalone HTML5/Chart.js interactive clinical dashboard web app      |
    +--------------------------------------------------------------------------+
```

---

## Repository File Manifest

The repository contains the modular Google Apps Script backend routines, data ingestion tools, and frontend dashboard components powering the spreadsheet:

| File Name | Component Type | Description |
| :--- | :---: | :--- |
| `Add New Row at Midnight.gs` | Scheduled Automation | Time-driven midnight cron job that initializes a new daily tracking row in `Tracking` with date indexing and baseline values. |
| `Drive Folder Indexer.gs` | Document Sync | Scans target Google Drive folders via Script Properties to index diagnostic PDFs and URLs into `Lab_Work` and `Discharge_Summaries`. |
| `Litter Robot Weight Import...gs` | Telemetry Ingestion | Parses exported weight event CSVs from the Whisker Litter-Robot 4 scale into `Weight_Log` and computes daily averages. |
| `Medication Logging.gs` | Clinical Handler | Manages daily therapeutic dose entries (Naraquin Phosphate Binder, Amlodipine, K-Gel, Fish Oil, Kidney Support Gold). |
| `Menu.gs` | Sheets UI Extension | Injects custom menu dropdowns into Google Sheets for manual sync triggers, report generation, and modal launchers. |
| `MobileTriggers.gs` | Event Pipeline | Event-driven triggers executing intake aggregations, schema validation, and urgent metric alerts on form submission. |
| `PetLibro Water Import.gs` | Telemetry Ingestion | Ingests RFID-tagged drinking events from the PetLibro Dockstream Smart Fountain, logging volume and timing into `Water_Log`. |
| `DarbepoetinLog.gs` | Cycle Management | Calculates administration intervals, cycle days, and scheduling alerts for subcutaneous Darbepoetin (21-day cycle). |
| `hydra slurry combine.gs` | Intake Math | Aggregates blended nutritional fluids, calculating net volume contributed by Hydra Care packets and wet food slurries. |
| `Dashboard_Script.gs` | Web Service Backend | Backend controller providing `doGet` web app hosting, modal dialog launchers, and clinical JSON endpoints (`getDashboardData`, `getProtocolData`). |
| `Dashboard.html` | Frontend Application | Standalone responsive HTML5/CSS/JavaScript dashboard featuring vital cards, fluid intake progress bars, and historical protocol accordions. |

---

## Physical Column Mappings

### 1. Master Daily Tracking (`Tracking`)

| Col Index | Header Field | Data Type | Clinical Target / Constraints | Description |
| :---: | :--- | :--- | :--- | :--- |
| **A** | `Date` | Date | `YYYY-MM-DD` | Master chronological index. |
| **B** | `Sub-Q (ml)` | Decimal | Standard: `100 mL` daily | Lactated Ringer's administered via gravity drip. |
| **C** | `Slurry & Hydracare (ml)` | Decimal | `>= 40 mL` daily | Syringe/bowl feeding (Weruva Phos-Focused Tuna + Hydra Care packet). |
| **D** | `Fountain (ml)` | Decimal | Dynamic | Cumulative intake logged from PetLibro RFID Dockstream fountain. |
| **E** | `Weight (lbs)` | Decimal | `0.00` – `20.00` | Calibrated dry weight or Litter-Robot 4 scale entry. |
| **F** | `K+ gel` | String | `1 tsp` daily | Oral potassium gluconate supplement to replenish renal loss. |
| **G** | `Fish Oil (ml)` | String | `1/4 tsp` daily | Omega-3 EPA/DHA fatty acids for renal inflammation support. |
| **H** | `Phosphate Binder` | String | `1 capsule` BID | Naraquin phosphate binder administered with meals. |
| **I** | `Kidney Support Gold (ml)`| Decimal | `5 drops (0.25 mL)` | Herbal renal support tincture. |
| **J** | `Prednisone` | String | Taper / Protocol | Oral anti-inflammatory / appetite protocol record. |
| **K** | `Darbopoetin` | Boolean | `TRUE` / `FALSE` | Erythropoiesis-stimulating agent administered (0.2 mL on 21-day cycle). |
| **L** | `Amlodipine` | String | `0.5 mL` q12h | Compounded blood pressure suspension (2.5 mg/mL). |
| **M** | `Total Intake (ml)` | Formula | `>= 200 mL` Floor | Calculated daily fluid volume: `SUM(Sub-Q + Slurry + Fountain)`. |
| **N** | `Other Medications Given` | String | Free-form | PRN therapeutics (e.g., Cerenia 8 mg, Mirtazapine transdermal). |
| **O** | `Notes` | String | Free-form | Clinical observations, uremic symptoms, or veterinary notes. |

### 2. Fountain Telemetry (`Water_Log`)

| Col Index | Header Field | Data Type | Description |
| :---: | :--- | :--- | :--- |
| **A** | `Date` | Date (`YYYY-MM-DD`) | Date of fountain interaction. |
| **B** | `Time` | Time (`HH:MM:SS`) | Timestamp of RFID collar tag detection. |
| **C** | `Amount (ml)` | Integer | Net weight difference converted to volume consumed. |
| **D** | `Daily Running Total (ml)`| Formula | Cumulative daily volume synced directly into `Tracking!D:D`. |

### 3. Scale Telemetry (`Weight_Log`)

| Col Index | Header Field | Data Type | Description |
| :---: | :--- | :--- | :--- |
| **A** | `Action` | String | Sensor event classification (`Weight recorded`). |
| **B** | `Date/Time` | Datetime | Timestamp when smart scale stabilized. |
| **C** | `Weight` | Decimal (lbs) | Patient mass recorded by Litter-Robot 4 sensors. |

### 4. Cyclical Therapeutic Control (`Darbopoetin_Log`)

| Col Index | Header Field | Data Type | Description |
| :---: | :--- | :--- | :--- |
| **A** | `Date` | Date (`YYYY-MM-DD`) | Logged date of check or administration. |
| **B** | `Time` | Time / Text | Time of injection or `MISSED DOSE` flag. |
| **C** | `Phase` | String | Protocol cycle description (`21-Day Cycle`). |
| **D** | `Cycle Day` | Integer (`1` – `21`) | Day counter within active dosing interval. |
| **E** | `Doses Given` | Integer | Cumulative lifetime dose counter. |

---

## Clinical Protocol & Observation Reference

### Daily Hydration Floor & Nutrition Target
* **Hydration Floor:** 200+ mL daily combined volume across Sub-Q gravity drip, slurry feedings, and RFID fountain.
* **Sub-Q Gravity Drip:** 100 mL Lactated Ringer's daily via gravity drip.
* **Slurry Base:** 40 mL Weruva Phos-Focused Tuna pâté + Hydra Care packet, 5 drops Kidney Support Gold, 1 tsp K-Gel, 1/4 tsp Fish Oil, and 1 capsule Naraquin Phosphate Binder twice daily.
* **Hypertension Management:** Amlodipine Besylate (2.5 mg/mL) 0.5 mL oral suspension every 12 hours.
* **RFID Micro-Feeder:** PetLibro RFID feeder loaded with Hill's Science Diet k/d (two 16 g automated portions daily).

### Palliative Symptom Index

| Observation / Symptom | Clinical Significance | Palliative Protocol / Action |
| :--- | :--- | :--- |
| **Mouth Pawing** | Signs of uremic nausea, gastric acid reflux, or "water brash". | Administer Maropitant (Cerenia 16 mg, 1/2 tablet) PRN with slurry. |
| **Teeth Grinding (Bruxing)** | Uremic gastritis, gastric hyperacidity, or oral discomfort. | Verify fluid intake; evaluate anti-nausea or acid-suppression support. |
| **Foul / Uremic Breath** | Elevated blood urea nitrogen (BUN) and circulating uremic toxins. | Increase fluid target toward 250 mL; inspect Sub-Q drip compliance. |
| **Neutral Breath** | Well-managed uremic state and effective metabolic clearance. | Continue baseline hydration and binder schedule. |
| **Marathon Drinking** | Active compensatory self-hydration at smart fountain. | Positive indicator; verify telemetry ingestion in `Water_Log`. |
| **Resting Respiration > 30 bpm** | Fluid overload warning or cardiac stress (Grade II/VI murmur). | Pause aggressive fluids; monitor resting sleeping respiratory rate. |
| **Left Hip Lesion** | Iliac crest localized irritation or over-grooming. | Inspect for scabbing, redness, or spreading alopecia; protect barrier. |

---

## Configuration & Deployment Setup

### 1. Script Properties Configuration
To avoid hardcoding sensitive Drive folder IDs and deployment endpoints in the source scripts, store them in Google Apps Script under **Project Settings > Script Properties**:

| Property (Key) | Example Value / Purpose |
| :--- | :--- |
| `FOLDER_ID_LAB_WORK` | Google Drive folder ID containing diagnostic lab reports and CBC/serum PDFs. |
| `FOLDER_ID_DISCHARGE_SUMMARIES`| Google Drive folder ID containing specialty clinic discharge summaries. |
| `WEB_APP_URL` | Deployed Google Apps Script `/exec` URL for loading the standalone dashboard. |

### 2. Scheduled Triggers
Configure the following time-driven and spreadsheet triggers in the Apps Script project dashboard:

* **`Add New Row at Midnight.gs`** -> Time-driven -> Day timer -> Midnight to 1:00 AM.
* **`Drive Folder Indexer.gs`** -> Time-driven -> Weekly / Daily timer.
* **`MobileTriggers.gs`** -> From spreadsheet -> On form submit.
* **`DarbepoetinLog.gs`** -> Time-driven -> Day timer -> 8:00 AM to 9:00 AM.
