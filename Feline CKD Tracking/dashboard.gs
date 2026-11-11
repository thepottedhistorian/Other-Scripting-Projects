/**
 * =========================================================================
 * SCRIPT: Zöe's Interactive Protocol & Clinical Care Dashboard Backend
 * AUTHOR: Rebecca Short
 * VERSION: 5.1 (Repaired Header Offset to Row 7, Exact Col Map & Fixed Math)
 * =========================================================================
 */

function doGet() {
  const htmlTemplate = HtmlService.createTemplateFromFile('Dashboard');
  return htmlTemplate.evaluate()
    .setTitle("Zöe's Clinical Dashboard & Protocol")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function showProtocolModal() {
  const webAppUrl = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') 
    || "YOUR_WEB_APP_EXEC_URL_HERE"; 
  
  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + webAppUrl + '", "_blank"); google.script.host.close();</script>'
  )
  .setWidth(350)
  .setHeight(100);

  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Dashboard...');
}

function getDashboardData(daysParam) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timeZone = ss.getSpreadsheetTimeZone() || "America/New_York";

    // 1. Fetch & Map Litter Box Entry Counts from 'Weight Log' Sheet
    const tripCountsByDate = {};
    const weightSheet = ss.getSheetByName("Weight Log");
    if (weightSheet) {
      const weightLastRow = weightSheet.getLastRow();
      if (weightLastRow >= 7) {
        const weightData = weightSheet.getRange(7, 2, (weightLastRow - 7) + 1, 1).getValues();
        for (let i = 0; i < weightData.length; i++) {
          const rawDateVal = weightData[i][0];
          if (!rawDateVal) continue;
          
          const parsedDate = new Date(rawDateVal);
          if (!isNaN(parsedDate.getTime()) && parsedDate.getFullYear() >= 2020) {
            const dateKey = Utilities.formatDate(parsedDate, timeZone, "MMM d");
            tripCountsByDate[dateKey] = (tripCountsByDate[dateKey] || 0) + 1;
          }
        }
      }
    }

    // 2. Fetch Darbepoetin Status Metrics from 'Darbepoetin Log' Sheet
    let darbepoetinStatus = {
      currentDay: 11,
      totalDays: 21,
      nextDoseDate: "Nov 18, 2026",
      totalDosesGiven: 6,
      lastDoseLogged: "Oct 28, 2026",
      phase: "21-Day Cycle"
    };

    const darboSheet = ss.getSheetByName("Darbepoetin Log") || ss.getSheetByName("Darbopoetin Log");
    if (darboSheet && darboSheet.getLastRow() >= 2) {
      const topRow = darboSheet.getRange(2, 1, 1, 5).getValues()[0];
      darbepoetinStatus.phase = topRow[2] || "21-Day Cycle";
      darbepoetinStatus.currentDay = Number(topRow[3]) || 11;
      darbepoetinStatus.totalDosesGiven = Number(topRow[4]) || 6;

      // Find last logged dose date in Column B
      const darboData = darboSheet.getRange(2, 1, darboSheet.getLastRow() - 1, 2).getValues();
      for (let i = 0; i < darboData.length; i++) {
        if (darboData[i][1] && darboData[i][1] !== "" && darboData[i][1] !== "MISSED DOSE") {
          const rawDoseDate = darboData[i][0];
          if (rawDoseDate) {
            const parsedDoseDate = new Date(rawDoseDate);
            darbepoetinStatus.lastDoseLogged = !isNaN(parsedDoseDate.getTime())
              ? Utilities.formatDate(parsedDoseDate, timeZone, "MMM d, yyyy")
              : String(rawDoseDate);
          }
          break;
        }
      }

      // Calculate Next Dose Date (Cycle counts from Day 1 to 21)
      const daysUntilNext = 21 - darbepoetinStatus.currentDay;
      const targetDate = new Date();
      targetDate.setDate(targetDate.getDate() + daysUntilNext);
      darbepoetinStatus.nextDoseDate = Utilities.formatDate(targetDate, timeZone, "MMM d, yyyy");
    }

    // 3. Fetch Clinical Tracking Vitals from 'Tracking' Sheet
    const sheet = ss.getSheetByName("Tracking");
    if (!sheet) return { labels: [], weights: [], totalFluids: [], subQ: [], slurry: [], fountain: [], litterBoxTrips: [], darbepoetin: darbepoetinStatus };

    const lastRow = sheet.getLastRow();
    const dataStartRow = 7; // Top data row (under Header Row 6)
    if (lastRow < dataStartRow) return { labels: [], weights: [], totalFluids: [], subQ: [], slurry: [], fountain: [], litterBoxTrips: [], darbepoetin: darbepoetinStatus };

    let days = (daysParam === undefined || daysParam === null) ? 30 : daysParam;
    let totalRowsAvailable = (lastRow - dataStartRow) + 1;
    let numRowsToFetch = totalRowsAvailable;
    
    if (String(days).toLowerCase() !== 'all' && !isNaN(Number(days))) {
      numRowsToFetch = Math.min(totalRowsAvailable, Number(days));
    }

    // Read Columns A through P (16 columns) starting at Row 7
    const data = sheet.getRange(dataStartRow, 1, numRowsToFetch, 16).getValues();

    const labels = [];
    const weights = [];
    const totalFluids = [];
    const subQ = [];
    const slurry = [];
    const fountain = [];

    // Loop backward through fetched slice (oldest to newest for charts)
    for (let i = data.length - 1; i >= 0; i--) {
      const row = data[i];
      const rawDate = row[0];
      
      if (!rawDate || rawDate === "") continue;

      const dateObj = new Date(rawDate);
      if (isNaN(dateObj.getTime()) || dateObj.getFullYear() < 2020) continue;

      const shortDate = Utilities.formatDate(dateObj, timeZone, "MMM d");

      labels.push(shortDate);
      subQ.push(Number(row[1]) || 0);
      slurry.push(Number(row[2]) || 0);
      fountain.push(Number(row[3]) || 0);
      
      const rawWeight = parseFloat(row[4]);
      weights.push(!isNaN(rawWeight) && rawWeight > 0 ? rawWeight : null);
      
      const total = Number(row[13]) || Number(row[11]) || (Number(row[1]) + Number(row[2]) + Number(row[3]));
      totalFluids.push(total);
    }

    const litterBoxTrips = labels.map(label => tripCountsByDate[label] || 0);

    return {
      labels: labels,
      weights: weights,
      totalFluids: totalFluids,
      subQ: subQ,
      slurry: slurry,
      fountain: fountain,
      litterBoxTrips: litterBoxTrips,
      darbepoetin: darbepoetinStatus
    };
  } catch (err) {
    Logger.log("Error in getDashboardData: " + err.toString());
    return { labels: [], weights: [], totalFluids: [], subQ: [], slurry: [], fountain: [], litterBoxTrips: [], darbepoetin: null, error: err.toString() };
  }
}

function getProtocolData() {
  return {
    active: {
      date: "November 2, 2026",
      vet: "Board-Certified Veterinary Internist (DACVIM)",
      medications: [
        "<b>Naraquin Phosphate Binder:</b> Give 1 capsule by mouth twice daily until otherwise instructed.",
        "<b>Amlodipine 2.5mg/mL liquid:</b> Give 0.5 mL by mouth every 12 hours until otherwise instructed. <i>(Maintained at 0.5 mL BID; discharge text note 'again increased' was a clerical typo)</i>.",
        "<b>Darbepoetin 25 mcg/mL:</b> Administer 0.2 mL under the skin every 21 days until otherwise instructed. <i>(Interval extended to 21 days as HCT stabilized)</i>.",
        "<b>Mirtazapine transdermal:</b> Apply 1 1/2 inch strip to alternating ear pinna every 24–48 hours or as instructed by your veterinarian. Discard 30 days after opening.",
        "<b>Maropitant 16mg tablet:</b> Give 1/2 tablet by mouth every 24 hours as needed or as directed (Anti-nausea medication).",
        "<b>Hydra Care Feline:</b> Give 1 packet by mouth every 24 hours until otherwise instructed to increase hydration.",
        "<b>Iron Dextran INJ 100mg/mL per mL:</b> Administered 0.5 mL once in L epaxial muscle.",
        "<b>Kidney Support Gold:</b> Put 5 drops on food every 24 hours as instructed by your veterinarian.",
        "<b>Potassium Gel:</b> Give 1 teaspoon in food every 24 hours as instructed by your veterinarian.",
        "<b>Fish oil:</b> Give 1/4 teaspoon in food every 24 hours as instructed by your veterinarian.",
        "<b>LRS:</b> Administer 100 mL under the skin every 24 hours as instructed by your veterinarian."
      ]
    },
    history: [
      {
        date: "October 5, 2026",
        vet: "Internal Medicine Specialist (DACVIM)",
        items: [
          "<b>Amlodipine 2.5mg/mL liquid:</b> Give 0.5 mL by mouth every 12 hours. <i>(Increased from 0.45 mL for continued blood pressure management)</i>.",
          "<b>Darbepoetin 25 mcg/mL:</b> Administer 0.2 mL under the skin every 14 days.",
          "<b>Continue:</b> Naraquin (1 cap BID), Mirtazapine, Maropitant (PRN), Hydra Care (1 packet), Kidney Support Gold, Potassium Gel, Fish Oil, LRS (100mL daily)."
        ]
      },
      {
        date: "September 16, 2026",
        vet: "Internal Medicine Specialist (DACVIM)",
        items: [
          "<b>Amlodipine 2.5mg/mL liquid:</b> Give 0.45 mL by mouth every 12 hours.",
          "<b>Darbepoetin 25 mcg/mL:</b> Administer 0.2 mL under the skin every 10 days.",
          "<b>Continue:</b> Naraquin (1 cap BID), Mirtazapine, Maropitant (PRN), Hydra Care (1 packet), Kidney Support Gold, Potassium Gel, Fish Oil, LRS (100mL daily).",
          "<b>Discontinued:</b> Varenzin, Prednisone."
        ]
      },
      {
        date: "August 26, 2026",
        vet: "Internal Medicine Specialist (DACVIM)",
        items: [
          "<b>Amlodipine 2.5mg/mL liquid:</b> Give 0.35 mL by mouth every 12 hours.",
          "<b>Darbepoetin 25 mcg/mL:</b> Administer 0.2 mL under the skin every 7 days.",
          "<b>Iron Dextran INJ 100mg/mL:</b> Administered 0.5 mL once in left epaxial muscle.",
          "<b>Continue:</b> Naraquin (1 cap BID), Mirtazapine, Maropitant (PRN), Hydra Care (1 packet), Kidney Support Gold, Potassium Gel, Fish Oil, LRS (100mL daily).",
          "<b>Discontinued:</b> Varenzin, Prednisone."
        ]
      },
      {
        date: "August 7, 2026",
        vet: "Internal Medicine Specialist",
        items: [
          "<b>Amlodipine Besylate Susp 2.5mg/ml:</b> Give 0.25ml by mouth every 12 hours until otherwise instructed.",
          "<b>Continue:</b> Kidney Support Gold, Naraquin, Potassium Gel, Fish Oil, LRS (100ml daily), Varenzin (0.9ml qd).",
          "<b>Prednisolone 10mg/ml Susp:</b> Continue weaning protocol."
        ]
      },
      {
        date: "July 20, 2026",
        vet: "Internal Medicine Specialist",
        items: [
          "<b>ADD:</b> Amlodipine Besylate 2.5mg - 1/4 tablet by mouth every 24 hours at night.",
          "<b>Continue:</b> Kidney Support Gold, Naraquin, Potassium Gel, Fish oil, LRS, Varenzin, Prednisone."
        ]
      },
      {
        date: "July 13, 2026",
        vet: "Primary Care Veterinarian",
        items: [
          "<b>ADD:</b> Prednisone 0.5ml 2x/day for 5 days then 0.5ml 1x/day for 10 days, then 0.5ml QOD until gone.",
          "<b>Continue:</b> 100ml Lactated Ringers SubQ daily, Naraquin Phosphate Binder, Varenzin CA1 0.8ml daily."
        ]
      },
      {
        date: "June 12, 2026",
        vet: "Primary Care Veterinarian",
        items: [
          "<b>ADD:</b> Veraflox Antibiotics 0.5ml for 10 days. 100ml Lactated Ringers SubQ daily, Naraquin Phosphate Binder, Varenzin CA1 0.8ml daily."
        ]
      },
      {
        date: "April 17, 2026",
        vet: "Primary Care Veterinarian",
        items: [
          "100ml Lactated Ringers SubQ daily, 1 tsp K+ gel daily, 1/4 tsp Fish Oil daily, Kidney Support Gold, Phosphate Binder."
        ]
      },
      {
        date: "April 7, 2026",
        vet: "Primary Care Veterinarian",
        items: [
          "100ml Lactated Ringers SubQ 2x/week, Naraquin Phosphate Binder (1 capsule twice daily), 1 tsp K+ gel daily, 1/4 tsp Fish Oil daily."
        ]
      }
    ]
  };
}
