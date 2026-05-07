/**
 * Creates a custom menu in Google Sheets.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('RS Clan Sync')
      .addItem('Full Sync (All Stages)', 'RunFullSync')
      .addSeparator()
      .addItem('Stage 1: Refresh Member List', 'Stage1_FetchRSList')
      .addItem('Stage 2: Resolve Missing IDs', 'Stage2_ResolveIDs')
      .addItem('Stage 3: Update Data', 'Stage3_FinalCombine')
      .addItem('Stage 4: Apply Formatting', 'ApplyClanFormatting')
      .addToUi();
  CreateScrollableLegend();
}

/**
 * Runs all three stages in sequence.
 */
function RunFullSync() {
  Stage1_FetchRSList();
  Stage2_ResolveIDs(); 
  Stage3_FinalCombine();
  ApplyClanFormatting();
}

// --- STAGE 1: OFFICIAL RS LIST & RUNEPIXELS CACHE ---
/**
 * Helper to ensure the master sheet is correctly named and retrieved.
 * If "Clan Master Sheet" is missing, it renames the active sheet.
 */
function getMasterSheet(ss) {
  let sheet = ss.getSheetByName("Clan Master Sheet");
  if (!sheet) {
    sheet = ss.getActiveSheet();
    sheet.setName("Clan Master Sheet");
  }
  return sheet;
}

// --- STAGE 1: OFFICIAL RS LIST & RUNEPIXELS CACHE ---
function Stage1_FetchRSList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getMasterSheet(ss);

  // Retrieve dynamic inputs from Row 1 (D1:F1 Merged)
  const clanName = sheet.getRange("B1").getValue().toString().trim();

  if (!clanName) {
    ss.toast("Error: Please ensure Clan Name is filled in Row 1.");
    return;
  }

  const clanId = getClanId(clanName);
  if (!clanId) return;

  const rsUrl = `https://secure.runescape.com/m=clan-hiscores/members_lite.ws?clanName=${encodeURIComponent(clanName)}`;
  const rpListUrl = `https://api.runepixels.com/clans/${clanId}/list`;

  // 1. Fetch Official RS List (The source of truth for current members)
  const rsRes = UrlFetchApp.fetch(rsUrl);
  const rsCsv = Utilities.parseCsv(rsRes.getContentText());

  // 2. Fetch RunePixels /list Data
  let rpLookup = {};
  try {
    const rpRes = UrlFetchApp.fetch(rpListUrl);
    const rpData = JSON.parse(rpRes.getContentText());
    rpData.forEach(p => {
      // Create a key using the same cleaning logic to ensure a match
      const keyName = p.name.replace(/[^A-Za-z0-9\s\-_]/g, " ").trim().replace(/[\s\-_]+/g, "-");
      rpLookup[keyName] = {
        id: p.playerID,
        joined: p.joinDate ? p.joinDate.split('T')[0] : "Unknown",
        activity: p.lastTimeActivity ? p.lastTimeActivity.split('T')[0] : "Unknown"
      };
    });
  } catch (e) {
    console.log("Error fetching RP list: " + e.message);
  }

  const newData = [];
  for (let i = 1; i < rsCsv.length; i++) {
    const rawName = rsCsv[i][0];
    // Clean names to remove diamonds and format for URL/Lookup
    const cleanName = rawName.replace(/[^A-Za-z0-9\s\-_]/g, " ").trim().replace(/[\s\-_]+/g, "-");
    
    if (cleanName) {
      const rpInfo = rpLookup[cleanName] || {};
      const joinDate = rpInfo.joined || "N/A";
      newData.push([
        cleanName,
        rsCsv[i][1],
        rsCsv[i][2], // Clan XP (from RS List)
        rpInfo.id || "",
        "Never", // Citadel Placeholder
        joinDate,
        calculateYYMMDD(joinDate),
        rpInfo.activity || "N/A"
      ]);
    }
  }
  
  // Clear existing data rows (A-H) to handle list shrinkage, preserving Row 1 Legend/Headers
  const lastRow = sheet.getLastRow();
  if (lastRow > 4) {
    sheet.getRange(5, 1, lastRow - 4, 8).clearContent();
  }
  if (newData.length > 0) {
    sheet.getRange(5, 1, newData.length, 8).setValues(newData);
  }
  
  FormatRSNColumn()
  ss.toast("Stage 1 Complete: Data synced from RunePixels /list.");
}

/**
 * Creates a static legend in the frozen top row.
 */
function CreateScrollableLegend() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getMasterSheet(ss);
  
  // 2. Preserve user inputs before clearing the UI area
  const existingClanName = sheet.getRange("B1").getValue();
  const existingResetTime = sheet.getRange("E2").getValue();
  const existingResetDay = sheet.getRange("G2").getValue();
  
  // 1. Clean slate: Ensure columns are visible and unmerged before we start building
  if (sheet.getMaxColumns() < 8) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 8 - sheet.getMaxColumns());
  }
  // Reset visibility and hide the technical ID column (D)
  sheet.showColumns(1, 8);
  sheet.hideColumns(4, 1); // Hide Column D (RunePixels ID)
  
  const uiArea = sheet.getRange(1, 1, 4, 8);
  uiArea.breakApart(); 
  uiArea.clearDataValidations();
  uiArea.clearContent();

  // --- ROW 1: CLAN INFO ---
  sheet.getRange("A1").setValue("Clan Name:").setFontWeight("bold").setHorizontalAlignment("right"); // Label
  const nameInput = sheet.getRange("B1"); // Input cell
  nameInput.setBackground("#fff2cc").setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  nameInput.setHorizontalAlignment("left").setFontWeight("italic"); // Align left for overflow
  nameInput.setValue(existingClanName || "Enter Clan Name");
  nameInput.setNote("Enter Clan Name here.");

  sheet.getRange("E1").setValue("Join Date based on RunePixels tracking data. Private RuneMetrics profile may cause discrepancies.").setFontStyle("italic").setFontSize(9).setHorizontalAlignment("left");

  // --- ROW 2: AUDIT CONTROLS ---
  // Date Input
  sheet.getRange("A2").setValue("Check Date:").setFontWeight("bold").setHorizontalAlignment("right");
  const dateInputCell = sheet.getRange("B2");
  dateInputCell.setBackground("#fff2cc").setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Time Input - Skipping Column C (Hidden)
  sheet.getRange("C2").setValue("Reset Time (UTC):").setFontWeight("bold").setHorizontalAlignment("right");
  const timeInputCell = sheet.getRange("E2");
  timeInputCell.setBackground("#fff2cc").setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  timeInputCell.setHorizontalAlignment("center");
  timeInputCell.setValue(existingResetTime || "00:00");

  // Reset Day Input
  sheet.getRange("F2").setValue("Reset Day:").setFontWeight("bold").setHorizontalAlignment("right");
  const resetDayInput = sheet.getRange("G2");
  resetDayInput.setBackground("#fff2cc").setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  resetDayInput.setHorizontalAlignment("center");
  resetDayInput.setValue(existingResetDay || "Wednesday");
  resetDayInput.setNote("Select the clan's weekly reset day.");

  // --- ROW 3: LEGEND ---
  const legendData = [["", "LEGEND:", "Missed Citadel", "", "Active Today", "7d+ Inactive", "14d+ Inactive", "21d+ Inactive"]];
  const legendRange = sheet.getRange(3, 1, 1, 8);
  legendRange.setValues(legendData).setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  sheet.getRange("B3").setFontWeight("italic");
  sheet.getRange("C3").setBackground("#ea9999");
  sheet.getRange("E3").setBackground("#b7e1cd"); // Active Today
  sheet.getRange("F3").setBackground("#f4cccc");
  sheet.getRange("G3").setBackground("#e06666");
  sheet.getRange("H3").setBackground("#990000").setFontColor("white");

  // --- ROW 4: HEADERS ---
  const headerData = [["RSN", "Rank", "Clan XP", "RunePixels ID", "Last Citadel Visit", "Join Date", "Time Since Joined", "Last Activity"]];
  sheet.getRange(4, 1, 1, 8).setValues(headerData);
  FormatHeaders();
  sheet.setFrozenRows(4); // Pins the legend and headers to the top

  // --- DATA VALIDATION ---
  // Time Validation
  const timeRule = SpreadsheetApp.newDataValidation()
    .requireTextMatchesPattern("^([01]?[0-9]|2[0-3]):[0-5][0-9]$")
    .setAllowInvalid(false)
    .setHelpText("Enter time in 24h format (e.g., 23:59).")
    .build();
  timeInputCell.setDataValidation(timeRule);

  // Day Validation
  const dayRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"])
    .setAllowInvalid(false)
    .setHelpText("Select the clan's reset day.")
    .build();
  resetDayInput.setDataValidation(dayRule);

  // Date Picker Validation
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate() 
    .setAllowInvalid(true) // Allow clearing the cell
    .setHelpText("Select a date for audit mode. Clear cell for current week tracking.")
    .build();
  dateInputCell.setDataValidation(dateRule);
  dateInputCell.setNote("Select a date here for audit mode. Clear cell for current week tracking.");

  // 4. Final UI Polish: Set specific column widths instead of auto-resizing
  sheet.setColumnWidth(1, 110); // RSN
  sheet.setColumnWidth(2, 95);  // Rank
  sheet.setColumnWidth(3, 125); // Clan XP
  sheet.setColumnWidth(4, 90);  // RunePixels ID (Hidden)
  sheet.setColumnWidth(5, 125); // Last Citadel Visit
  sheet.setColumnWidth(6, 90);  // Join Date
  sheet.setColumnWidth(7, 125); // Time Since Joined
  sheet.setColumnWidth(8, 90);  // Last Activity

  // 5. Final visibility check and UI commitment
  SpreadsheetApp.flush();
  sheet.hideColumns(4, 1);
}
/**
 * Center justifies the header row (Row 1).
 */
function FormatHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getMasterSheet(ss);
  
  // Select the entire first row where your headers are
  const headerRange = sheet.getRange(4, 1, 1, 8);
  
  // Set alignment to Center and make it bold for a cleaner look
  headerRange.setHorizontalAlignment("center");
  headerRange.setFontWeight("bold");
  
  ss.toast("Headers centered and bolded.");
}
/**
 * Formats Column A to be right-aligned with a 2-space indent.
 */
function FormatRSNColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getMasterSheet(ss);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 5) return;

  const rsnRange = sheet.getRange(5, 1, lastRow - 4, 1);
  const rankRange = sheet.getRange(5, 2, lastRow - 4, 1);

  // 1. Set alignment
  rsnRange.setHorizontalAlignment("right");
  rankRange.setHorizontalAlignment("center");

  // 2. Apply Custom Number Format: Text followed by two spaces
  // The "@" represents the text, and the spaces after it create the indent
  rsnRange.setNumberFormat('@"  "');
  
  ss.toast("RSN Column formatted with right-indent.");
}

// --- STAGE 2: RESOLVE MISSING IDs (Loop with Time Limit) ---
function Stage2_ResolveIDs() {
  const startTime = new Date().getTime();
  const maxRuntime = 330000; 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getMasterSheet(ss);
  const lastRow = sheet.getLastRow();
  ss.toast('Resolving RunePixels IDs, this step may take a few minutes.')
  
  if (lastRow < 5) return;
  const range = sheet.getRange(5, 1, lastRow - 4, 4);
  const data = range.getValues();
  let count = 0;

  for (let i = 0; i < data.length; i++) {
    if (new Date().getTime() - startTime > maxRuntime) break; //

    const rsn = data[i][0];
    const currentId = data[i][3];

    // Only check if ID is blank and we have a valid RSN
    if (!currentId && rsn && rsn !== "") { 
      const id = getPlayerId(rsn);
      if (id) {
        data[i][3] = id;
        count++;
        if (count % 10 === 0) range.setValues(data);
      }
    }
  }
  range.setValues(data);
  ss.toast(`Stage 2: Resolved ${count} additional IDs.`);
}

/**
 * STAGE 3: POST REQUEST SYNC
 * Uses the POST method discovered in the network trace to bypass 405 errors.
 */
function Stage3_FinalCombine() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getMasterSheet(ss);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 5) return;

  const clanName = sheet.getRange("B1").getValue().toString().trim();
  const clanId = getClanId(clanName);
  if (!clanId) return;

  const listUrl = `https://api.runepixels.com/clans/${clanId}/list`;
  const citadelUrl = `https://api.runepixels.com/clans/${clanId}/players/citadel`;

  // 1. Prepare Headers and Payload based on your network trace
  const options = {
    "method": "post",
    "contentType": "application/json",
    "headers": {
      "Accept": "application/json, text/plain, */*",
      "Origin": "https://runepixels.com",
      "Referer": "https://runepixels.com/"
    },
    "payload": JSON.stringify({ "search": "", "rank": -1 }), // Standard body for this endpoint
    "muteHttpExceptions": true
  };

  // 2. Fetch the Data
  let citadelData = {};
  let clanListLookup = {};

  try {
    // Fetch Clan List via POST
    const listRes = UrlFetchApp.fetch(listUrl, options);
    if (listRes.getResponseCode() === 200) {
      const listData = JSON.parse(listRes.getContentText());
      listData.forEach(p => {
        clanListLookup[p.playerID] = {
          joined: p.joinDate ? p.joinDate.split('T')[0] : "N/A",
          activity: p.lastTimeActivity ? p.lastTimeActivity.split('T')[0] : "N/A"
        };
      });
    } else {
      console.log("List fetch failed: " + listRes.getResponseCode());
    }

    // Fetch Citadel (Usually remains a GET)
    const citRes = UrlFetchApp.fetch(citadelUrl, {"muteHttpExceptions": true});
    if (citRes.getResponseCode() === 200) citadelData = JSON.parse(citRes.getContentText());

  } catch (e) {
    ss.toast("Error during fetch: " + e.message);
  }

  // PRE-PROCESS CITADEL DATA (Optimization: Parse strings once, not inside the loop)
  const manualDate = sheet.getRange("B2").getValue();
  const resetTimeInput = sheet.getRange("E2").getValue();
  const useFilter = (manualDate instanceof Date && !isNaN(manualDate));
  
  // Parse Reset Time for automatic tracking
  let resetHour = 0, resetMin = 0;
  if (resetTimeInput instanceof Date) {
    resetHour = resetTimeInput.getHours();
    resetMin = resetTimeInput.getMinutes();
  } else if (typeof resetTimeInput === "string" && resetTimeInput.includes(":")) {
    const parts = resetTimeInput.split(":");
    resetHour = parseInt(parts[0], 10);
    resetMin = parseInt(parts[1], 10);
  }

  // Define resetTimeStr for use in the filterCutoff timestamp
  const resetTimeStr = Utilities.formatString('%02d:%02d', resetHour, resetMin);

  const filterCutoff = useFilter ? 
    Utilities.formatDate(manualDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") + " " + resetTimeStr : 
    null;

  const citadelLookup = {};
  for (const pid in citadelData) {
    try {
      const history = JSON.parse(citadelData[pid]);
      if (Array.isArray(history) && history.length > 0) {
        if (filterCutoff) {
          // Find the most recent visit that occurred on or before the audit timestamp
          for (let j = 0; j < history.length; j++) {
            if (history[j] <= filterCutoff) {
              citadelLookup[pid] = history[j].split(' ')[0];
              break;
            }
          }
        } else {
          // Store the absolute most recent visit
          citadelLookup[pid] = history[0].split(' ')[0];
        }
      }
    } catch (e) { /* Skip malformed data */ }
  }

  // 3. Update the Master Sheet
  const range = sheet.getRange(5, 1, lastRow - 4, 8);
  const values = range.getValues();
  const output = [];

  for (let i = 0; i < values.length; i++) {
    const rawId = values[i][3]; // Column D
    let id = null;
    if (rawId !== null && rawId !== "") {
      const numId = Number(rawId);
      if (!isNaN(numId)) id = numId;
    }
    let citadelDate = "Never";
    // Fallback to existing sheet data from Columns F (index 5) and H (index 7)
    let joinDate = values[i][5] || "N/A";
    let lastActivity = values[i][7] || "N/A";

    if (id) {
      citadelDate = citadelLookup[id] || "Never";
      if (clanListLookup[id]) {
        joinDate = clanListLookup[id].joined;
        lastActivity = clanListLookup[id].activity;
      }
    }
    
    // Populate E (Citadel), F (Join Date), G (Time Since Joined), and H (Last Activity)
    output.push([citadelDate, joinDate, calculateYYMMDD(joinDate), lastActivity]);
  }

  // 4. Write back to E, F, G and H
  sheet.getRange(5, 5, output.length, 4).setValues(output);
  FormatCitadelColumn()
  ss.toast("Success! Columns E-H updated using POST.");
}

/**
 * Formats Column D to be left-aligned.
 * You can add this to your main formatting function.
 */
function FormatCitadelColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getMasterSheet(ss);
  const lastRow = sheet.getLastRow();
  
  if (lastRow < 5) return;

  const citadelRange = sheet.getRange(5, 5, lastRow - 4, 1);
  const joinRange = sheet.getRange(5, 6, lastRow - 4, 1);
  const timeSinceRange = sheet.getRange(5, 7, lastRow - 4, 1);
  const activityRange = sheet.getRange(5, 8, lastRow - 4, 1);

  // Set alignment and Date Picker for Column D
  citadelRange.setHorizontalAlignment("left")
    .setNumberFormat("yyyy-mm-dd");
  
  // Data Validation for Column D
  const manualDate = sheet.getRange("B2").getValue();
  let validationBuilder = SpreadsheetApp.newDataValidation().requireDate();
  
  // If an audit date is set, restrict the date picker/input to on or before that date
  if (manualDate instanceof Date && !isNaN(manualDate)) {
    validationBuilder.requireDateOnOrBefore(manualDate);
  }

  const citadelRule = validationBuilder.setAllowInvalid(true).build();
  citadelRange.setDataValidation(citadelRule);
  
  // Set alignment to Right
  joinRange.setHorizontalAlignment("right");
  timeSinceRange.setHorizontalAlignment("right");
  activityRange.setHorizontalAlignment("right")

  // Ensure Column C remains hidden after formatting updates
  // Ensure technical/inaccurate columns remain hidden
  sheet.hideColumns(4, 1);
  SpreadsheetApp.flush();
  
  ss.toast("Column alignments set.");
}

/**
 * STAGE 4: APPLY CONDITIONAL FORMATTING
 * Applies conditional formatting to clan data.
 */
function ApplyClanFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getMasterSheet(ss);
  const lastRow = sheet.getLastRow();
  if (lastRow < 5) return;

  sheet.clearConditionalFormatRules();
  const rules = [];

  // --- COLUMN E: CITADEL (RED IF BEFORE LAST FRIDAY) ---
  const manualDate = sheet.getRange("B2").getValue();
  const resetTimeInput = sheet.getRange("E2").getValue();
  const resetDayStr = sheet.getRange("G2").getValue();
  let targetDate;

  // Parse Reset Time
  let resetHour = 0, resetMin = 0;
  if (resetTimeInput instanceof Date) {
    resetHour = resetTimeInput.getHours();
    resetMin = resetTimeInput.getMinutes();
  } else if (typeof resetTimeInput === "string" && resetTimeInput.includes(":")) {
    const parts = resetTimeInput.split(":");
    resetHour = parseInt(parts[0], 10);
    resetMin = parseInt(parts[1], 10);
  }

  const dayMap = {
    "Sunday": 0, "Monday": 1, "Tuesday": 2, "Wednesday": 3,
    "Thursday": 4, "Friday": 5, "Saturday": 6
  };
  const resetDayNum = dayMap[resetDayStr] !== undefined ? dayMap[resetDayStr] : 5;

  if (manualDate instanceof Date && !isNaN(manualDate)) {
    // Audit Mode: To find who missed the week ending on [B2], we flag visits
    // that occurred BEFORE that week started (i.e., 7 days prior to B2).
    targetDate = new Date(manualDate.getTime());
    targetDate.setDate(targetDate.getDate() - 7);
  } else {
    // Auto-calculate the current cycle based on reset time
    const now = new Date();
    const dayOfWeek = now.getDay();
    let daysToSubtract = (dayOfWeek - resetDayNum + 7) % 7;

    // If today is the reset day, check if the reset has actually happened yet
    if (dayOfWeek === resetDayNum) {
      const resetTimeToday = new Date(now.getTime());
      resetTimeToday.setHours(resetHour, resetMin, 0, 0); 
      if (now < resetTimeToday) daysToSubtract = 7; // Reset hasn't happened; target last week
    }

    targetDate = new Date();
    targetDate.setDate(now.getDate() - daysToSubtract);
  }

  const fridayIso = Utilities.formatDate(targetDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");

  const rangeE = sheet.getRange(5, 5, lastRow - 4, 1);
  const ruleCitadel = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=OR($E5="Never", AND(ISDATE($E5), $E5 < DATEVALUE("${fridayIso}")))`)
    .setBackground("#ea9999") // Light Red
    .setRanges([rangeE])
    .build();
  rules.push(ruleCitadel);

  // --- COLUMN H: ACTIVITY HEATMAP (DARKER = OLDER) ---
  const rangeH = sheet.getRange(5, 8, lastRow - 4, 1);

  // 0. Active Today (Light Green)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H5 = TODAY()')
    .setBackground("#b7e1cd") // Light Green
    .setRanges([rangeH])
    .build());

  // 1. Darkest Red (Over 21 days)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$H5 < (TODAY() - 21)')
    .setBackground("#990000")
    .setFontColor("#ffffff")
    .setRanges([rangeH])
    .build());

  // 2. Medium Red (14 - 21 days)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($H5 < (TODAY() - 14), $H5 >= (TODAY() - 21))')
    .setBackground("#e06666")
    .setRanges([rangeH])
    .build());

  // 3. Light Red (7 - 14 days)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($H5 < (TODAY() - 7), $H5 >= (TODAY() - 14))')
    .setBackground("#f4cccc")
    .setRanges([rangeH])
    .build());

  sheet.setConditionalFormatRules(rules);

  // Add mouseover notes for Column H (Last Activity) to show "days ago"
  const activityValues = rangeH.getValues();
  const activityNotes = activityValues.map(row => [calculateDaysAgo(row[0]) || ""]);
  rangeH.setNotes(activityNotes);

  // Final UI Polish: Ensure only necessary columns are visible
  sheet.showColumns(1, 8);
  sheet.getRange(5, 3, lastRow - 4, 1).setNumberFormat("#,##0");
  SpreadsheetApp.flush();

  // Final UI Polish: Set specific column widths
  sheet.setColumnWidth(1, 110); // RSN
  sheet.setColumnWidth(2, 95);  // Rank
  sheet.setColumnWidth(3, 125); // Clan XP
  sheet.setColumnWidth(4, 90);  // RunePixels ID (Hidden)
  sheet.setColumnWidth(5, 125); // Last Citadel Visit
  sheet.setColumnWidth(6, 90);  // Join Date
  sheet.setColumnWidth(7, 125); // Time Since Joined
  sheet.setColumnWidth(8, 90);  // Last Activity

  // Re-hide technical data
  // Re-hide technical and ignored data
  sheet.hideColumns(4, 1);
  SpreadsheetApp.flush();

  ss.toast("Success: Citadel & Activity alerts applied!");
}

function getPlayerId(rsn) {
  const url = `https://api.runepixels.com/players/${encodeURIComponent(rsn)}`;
  try {
    const res = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()).id;
  } catch (e) { return null; }
  return null;
}

/**
 * Fetches the Clan ID from RunePixels using the Clan Name.
 */
function getClanId(clanName) {
  if (!clanName) return null;
  const formattedClanName = clanName.toString().trim().toLowerCase().replace(/\s/g, '-');
  const url = `https://api.runepixels.com/clans/${encodeURIComponent(formattedClanName)}`;
  
  try {
    const res = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    if (res.getResponseCode() === 200) {
      const data = JSON.parse(res.getContentText());
      return data.id || null;
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Error fetching Clan ID: ${res.getResponseCode()}`);
    }
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error finding Clan ID: ${e.message}`);
  }
  return null;
}

/**
 * Calculates the time difference in YY MM DD format between a date and today.
 */
function calculateYYMMDD(dateStr) {
  if (!dateStr || dateStr === "N/A" || dateStr === "Unknown" || dateStr === "") return "N/A";
  const start = new Date(dateStr);
  const today = new Date();
  if (isNaN(start.getTime())) return "N/A";
  
  let years = today.getFullYear() - start.getFullYear();
  let months = today.getMonth() - start.getMonth();
  let days = today.getDate() - start.getDate();
  
  if (days < 0) {
    months--;
    const lastDayPrevMonth = new Date(today.getFullYear(), today.getMonth(), 0).getDate();
    days += lastDayPrevMonth;
  }
  if (months < 0) {
    years--;
    months += 12;
  }
  
  return `${years}y ${months}m ${days}d`;
}

/**
 * Calculates how many days ago a date was from today.
 */
function calculateDaysAgo(dateStr) {
  if (!dateStr || dateStr === "N/A" || dateStr === "Unknown" || dateStr === "") return null;
  const activityDate = new Date(dateStr);
  if (isNaN(activityDate.getTime())) return null;
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  activityDate.setHours(0, 0, 0, 0);
  
  const diffTime = today - activityDate;
  const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
  
  if (diffDays < 0) return "Future date"; 
  if (diffDays === 0) return "Active today";
  if (diffDays === 1) return "1 day ago";
  return `${diffDays} days ago`;
}
