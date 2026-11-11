# Other Scripting Projects

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-ES6+-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Python](https://img.shields.io/badge/Python-3.x-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Google Workspace](https://img.shields.io/badge/Platform-Google%20Workspace-34A853?style=for-the-badge&logo=google&logoColor=white)
![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg?style=for-the-badge)
![License](https://img.shields.io/badge/License-All%20Rights%20Reserved-red?style=for-the-badge)

A centralized monorepo of standalone automation engines, Google Apps Script utilities, data ingestion pipelines, and custom scripting tools designed for practical workflow optimization, telemetry processing, and personal productivity.

---

## Repository Structure

```
Other-Scripting-Projects/
├── Feline CKD Tracking/
│   ├── README.md                          # Clinical architecture & hardware docs
│   ├── Add New Row at Midnight.gs         # Chronological row initializer
│   ├── Drive Folder Indexer.gs            # Diagnostic PDF drive synchronizer
│   ├── Litter Robot Weight Import...gs    # Whisker scale CSV parser
│   ├── Medication Logging.gs              # Dose tracker & binder logic
│   ├── Menu.gs                            # Custom Sheets UI dropdowns
│   ├── MobileTriggers.gs                  # Form-submit execution pipeline
│   ├── PetLibro Water Import.gs           # Smart fountain volume ingest
│   ├── charts.gs                          # Dynamic Sheets chart engines
│   ├── retrofiller.gs                     # Historical data reconciliation
│   ├── DarbepoetinLog.gs                  # Recombinant ESA 21-day cycle tracker
│   ├── hydra slurry combine.gs            # Intake aggregation calculator
│   ├── Dashboard_Script.gs                # doGet backend & JSON API endpoints
│   └── Dashboard.html                     # HTML5/Chart.js palliative care dashboard
│
└── README.md                              # Root repository index
```

---

## Projects Directory

### 🐾 [Feline CKD Tracking](./Feline%20CKD%20Tracking/)
* **Stack:** Google Apps Script, Google Sheets, Chart.js, HTML5/CSS, PetLibro RFID, Whisker Litter-Robot 4
* **Description:** An automated palliative veterinary telemetry suite engineered for Stage IV Chronic Kidney Disease care. Ingests RFID fountain drink logs, weight scale telemetry, multi-route fluid intake, and historical blood panels into a unified relational spreadsheet and web app dashboard.
* **Key Features:**
  * Automated 200+ mL daily hydration threshold monitoring and alerts.
  * RFID-tagged fountain bout logging and smart litter scale weight parsing.
  * 21-day cyclic medication scheduling for subcutaneous Darbepoetin.
  * Document indexing pipeline for specialist discharge summaries and serum lab reports.

---

## Automation Architecture & Standards

All sub-projects in this repository adhere to standard development guidelines:

* **Decoupled Configuration:** Sensitive IDs, private folder strings, API secrets, and `/exec` deployment endpoints are decoupled from source files using `PropertiesService.getScriptProperties()` or environment configurations.
* **Idempotent Executions:** Scheduled crons and form-trigger pipelines validate boundary conditions prior to appending rows or recalculating running balances.
* **Defensive Schemas:** Scripts incorporate header-offset protection and type checking to prevent silent failures during batch ingestion.

---

## Copyright & Intellectual Property

Copyright © 2026 Rebecca Short. All rights reserved.

The source code, custom schemas, dashboard markup, and clinical architecture contained within this repository are proprietary and created for private, non-commercial application. Unauthorized reproduction, redistribution, adaptation, or public hosting of these files—in whole or in part—without explicit written permission is strictly prohibited.
