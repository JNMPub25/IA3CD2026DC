/**
 * District 3 Convention — Google Apps Script
 * Receives ballot submissions from ballot.html and writes them to Google Sheets.
 *
 * SETUP INSTRUCTIONS (do this once before the convention):
 * ─────────────────────────────────────────────────────────
 * 1. Open Google Sheets → create a new spreadsheet called "District 3 Ballots"
 *    (create a second one called "District 3 Ballots — TEST" for testing)
 *
 * 2. In that spreadsheet: Extensions → Apps Script
 *
 * 3. Delete all default code and paste this entire file.
 *
 * 4. Update SPREADSHEET_ID below:
 *    - Open the spreadsheet in your browser
 *    - Copy the long ID from the URL:
 *      https://docs.google.com/spreadsheets/d/ >>>THIS PART<<< /edit
 *
 * 5. Click Deploy → New deployment
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 *    - Click Deploy, authorize when prompted
 *    - Copy the deployment URL — paste it into ballot.html CONFIG.APPS_SCRIPT_URL
 *
 * 6. Repeat steps 1–5 for the TEST spreadsheet,
 *    and paste that URL into CONFIG.TEST_APPS_SCRIPT_URL in ballot.html
 *
 * RE-DEPLOYING AFTER CHANGES:
 *    Deploy → Manage deployments → Edit (pencil) → Version: New version → Deploy
 *    (The URL stays the same after the first deployment as long as you manage it
 *     through "Manage deployments" rather than creating a new one.)
 */

// ── UPDATE THIS ──────────────────────────────────────────────────────────
const SPREADSHEET_ID = "PASTE_YOUR_SPREADSHEET_ID_HERE";

// Name of the control tab used by the Option C timer
const CONTROL_TAB = "Control";

// Google Drive folder ID for "3rd CD Convention"
// (from the URL when you open that folder in Drive)
const CONVENTION_FOLDER_ID = "1qapYmQN_oicEyYKbuvU9zq6eP3QrNwKK";
// ──────────────────────────────────────────────────────────────────────────

/**
 * ONE-TIME SETUP: Move this spreadsheet into the "3rd CD Convention" folder.
 *
 * Run this manually from the Apps Script editor (not via deployment):
 *   1. Open the Apps Script editor for this spreadsheet
 *   2. Select "moveToConventionFolder" from the function dropdown
 *   3. Click Run
 *   4. Repeat in the TEST spreadsheet's Apps Script editor
 *
 * You only need to do this once per spreadsheet.
 */
function moveToConventionFolder() {
  const file   = DriveApp.getFileById(SPREADSHEET_ID);
  const folder = DriveApp.getFolderById(CONVENTION_FOLDER_ID);
  file.moveTo(folder);
  Logger.log("Moved spreadsheet " + SPREADSHEET_ID + " to 3rd CD Convention folder.");
}

// ── ELECTION NAME LOOKUP ─────────────────────────────────────────────────

/**
 * Maps election keys to human-readable names for sheet tabs and Setup tab.
 */
const ELECTION_LABELS = {
  "scc-w":   "SCC Women",
  "scc-m":   "SCC Men-NonBinary",
  "dei":     "DEI Chair",
  "scc-com": "Convention Committee",
};

function getSheetName(electionKey) {
  return ELECTION_LABELS[electionKey] || electionKey;
}

// ── CONTROL TAB HELPERS (Option C timer) ─────────────────────────────────

/**
 * Returns the Control tab, creating it if it doesn't exist.
 * Row 1 = header; Row 2 = single data row (overwritten on every action).
 * Columns: Status | OpenedAt | Duration | FloorCount
 */
function ensureControlTab(ss) {
  let sheet = ss.getSheetByName(CONTROL_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(CONTROL_TAB);
    sheet.appendRow(["Status", "OpenedAt", "Duration", "FloorCount"]);
    sheet.appendRow(["standby", "", "", ""]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 4).setFontWeight("bold");
  }
  return sheet;
}

function getControlRow(sheet) {
  const v = sheet.getRange(2, 1, 1, 4).getValues()[0];
  return {
    status:     String(v[0] || "standby"),
    openedAt:   String(v[1] || ""),
    duration:   Number(v[2] || 0),
    floorCount: Number(v[3] || 0),
  };
}

function setControlRow(sheet, status, openedAt, duration, floorCount) {
  sheet.getRange(2, 1, 1, 4).setValues([[status, openedAt, duration, floorCount]]);
}

// ── SETUP TAB HELPERS (election metadata) ────────────────────────────────

/**
 * Creates or updates a "[Election Name] Setup" tab with election metadata.
 * Layout is key-value pairs in two columns (Field | Value).
 *
 * Fields:
 *   Election Name, Election Key, Seats, Candidates, Ballot URL,
 *   Admin URL, Assisted URL, Start Time, End Time, Duration (min),
 *   Floor Count, Inactive Credentials
 *
 * Called by:
 *   - "setupElection" action (admin UI pushes setup before voting)
 *   - "open" action (auto-records start time)
 *   - "close" action (auto-records end time)
 *   - "updateSetup" action (manual override of any field)
 */
function getSetupTabName(electionKey) {
  return getSheetName(electionKey) + " Setup";
}

function ensureSetupTab(ss, electionKey) {
  const tabName = getSetupTabName(electionKey);
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    // Build the key-value structure
    const fields = [
      ["Field", "Value"],
      ["Election Name", getSheetName(electionKey)],
      ["Election Key", electionKey],
      ["Seats", ""],
      ["Ballot Type", ""],
      ["Candidates", ""],
      ["Candidates Live At", ""],
      ["Ballot URL", ""],
      ["Admin URL", ""],
      ["Assisted URL", ""],
      ["Start Time", ""],
      ["End Time", ""],
      ["Duration (min)", ""],
      ["Floor Count", ""],
      ["Inactive Credentials", ""],
    ];
    sheet.getRange(1, 1, fields.length, 2).setValues(fields);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 2).setFontWeight("bold");
    // Style the field column
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 500);
    sheet.getRange(2, 1, fields.length - 1, 1).setFontWeight("bold");
  }
  return sheet;
}

/**
 * Sets a single field in the Setup tab by finding its row.
 * fieldName must match exactly one of the Field column values.
 */
function setSetupField(sheet, fieldName, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === fieldName) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
}

/**
 * Gets a single field value from the Setup tab.
 */
function getSetupField(sheet, fieldName) {
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === fieldName) {
      return data[i][1];
    }
  }
  return "";
}

/**
 * Populates the Setup tab with full election metadata from the admin UI.
 */
function populateSetupTab(ss, data) {
  const electionKey = data.electionKey;
  const sheet = ensureSetupTab(ss, electionKey);

  if (data.seats !== undefined)        setSetupField(sheet, "Seats", data.seats);
  if (data.candidates)                 setSetupField(sheet, "Candidates", data.candidates);
  if (data.ballotUrl)                  setSetupField(sheet, "Ballot URL", data.ballotUrl);
  if (data.adminUrl)                   setSetupField(sheet, "Admin URL", data.adminUrl);
  if (data.assistedUrl)                setSetupField(sheet, "Assisted URL", data.assistedUrl);
  if (data.duration !== undefined)     setSetupField(sheet, "Duration (min)", data.duration);
  if (data.inactiveCredentials)        setSetupField(sheet, "Inactive Credentials", data.inactiveCredentials);
}

// ── GET: status polling endpoint ─────────────────────────────────────────

/**
 * Handles GET requests from ballot.html (status polling) and the setup script
 * (connectivity test).
 *
 * Query params:
 *   ?election=scc-w   → also returns voteCount for that election tab
 *
 * Response:
 *   { status: "standby"|"open"|"closed", openedAt, duration, floorCount,
 *     voteCount, message }
 */
function doGet(e) {
  const electionKey = (e && e.parameter && e.parameter.election)
    ? e.parameter.election : "";

  try {
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ctrl = ensureControlTab(ss);
    const row  = getControlRow(ctrl);

    let voteCount = 0;
    if (electionKey) {
      const sheetName = getSheetName(electionKey);
      const elSheet   = ss.getSheetByName(sheetName);
      if (elSheet) voteCount = Math.max(0, elSheet.getLastRow() - 1);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        status:     row.status,
        openedAt:   row.openedAt,
        duration:   row.duration,
        floorCount: row.floorCount,
        voteCount:  voteCount,
        message:    "Ballot receiver is live.",
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── POST: ballot submissions + admin control actions ─────────────────────

/**
 * Receives POST requests from ballot.html and the admin dashboard.
 *
 * Request types:
 *
 *  1. Admin control (action field present):
 *     { action: "open",  openedAt, duration, floorCount, electionKey }
 *       → sets Control tab to "open", records start time in Setup tab
 *     { action: "close", electionKey }
 *       → sets Control tab to "closed", records end time in Setup tab
 *
 *  2. Setup metadata (from admin dashboard):
 *     { action: "setupElection", electionKey, seats, candidates,
 *       ballotUrl, adminUrl, assistedUrl, duration, inactiveCredentials }
 *       → creates/updates the "[Election] Setup" tab
 *
 *  3. Manual timing override:
 *     { action: "updateSetup", electionKey, field, value }
 *       → updates a single field in the Setup tab
 *
 *  4. Ballot submission (no action field):
 *     Each election gets its own tab.
 *     If a delegate re-submits, the previous row is overwritten (last vote wins).
 */
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);

    // ── Admin: open voting ───────────────────────────────────────────────
    if (data.action === "open") {
      const ctrl = ensureControlTab(ss);
      setControlRow(ctrl, "open", data.openedAt || "", data.duration || 0, data.floorCount || 0);

      // Auto-record start time in Setup tab
      if (data.electionKey) {
        const setupSheet = ensureSetupTab(ss, data.electionKey);
        setSetupField(setupSheet, "Start Time", data.openedAt || now_());
        if (data.duration)   setSetupField(setupSheet, "Duration (min)", data.duration);
        if (data.floorCount) setSetupField(setupSheet, "Floor Count", data.floorCount);
      }

      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Admin: close voting ──────────────────────────────────────────────
    if (data.action === "close") {
      const ctrl = ensureControlTab(ss);
      const row  = getControlRow(ctrl);
      setControlRow(ctrl, "closed", row.openedAt, row.duration, row.floorCount);

      // Auto-record end time in Setup tab
      if (data.electionKey) {
        const setupSheet = ensureSetupTab(ss, data.electionKey);
        setSetupField(setupSheet, "End Time", now_());
      }

      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Setup election metadata (from admin dashboard) ───────────────────
    if (data.action === "setupElection") {
      if (!data.electionKey) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: "error", message: "electionKey required" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      populateSetupTab(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok", message: "Setup tab created/updated" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Manual update to Setup tab ───────────────────────────────────────
    if (data.action === "updateSetup") {
      if (!data.electionKey || !data.field) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: "error", message: "electionKey and field required" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const setupSheet = ensureSetupTab(ss, data.electionKey);
      setSetupField(setupSheet, data.field, data.value || "");
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Export ballot data for tabulation (from admin dashboard) ────────
    if (data.action === "exportBallots") {
      if (!data.electionKey) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: "error", message: "electionKey required" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const sheetName = getSheetName(data.electionKey);
      const elSheet = ss.getSheetByName(sheetName);
      if (!elSheet || elSheet.getLastRow() < 2) {
        return ContentService
          .createTextOutput(JSON.stringify({ status: "ok", rows: [], headers: [] }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const allData = elSheet.getDataRange().getValues();
      const headers = allData[0].map(h => String(h).trim());
      const rows = [];
      for (let i = 1; i < allData.length; i++) {
        const obj = {};
        for (let j = 0; j < headers.length; j++) {
          obj[headers[j]] = String(allData[i][j] || "");
        }
        rows.push(obj);
      }
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok", headers: headers, rows: rows }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Ballot submission ────────────────────────────────────────────────
    const sheetName = data.ballotType === "runoff"
      ? "Runoff " + data.electionKey
      : getSheetName(data.electionKey);
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      writeHeader(sheet, data);
    }

    const isOverwrite = upsertRow(sheet, data);
    appendAuditLog(ss, data, isOverwrite);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ── TIMESTAMP HELPER ─────────────────────────────────────────────────────

function now_() {
  return Utilities.formatDate(new Date(), "America/Chicago", "MM/dd/yyyy hh:mm:ss a");
}

// ── HEADER / UPSERT HELPERS ──────────────────────────────────────────────

/**
 * Write header row based on the first submission received.
 * For ranked ballots: one column per candidate letter.
 * For slate ballots:  a single "Slate Vote" column.
 */
function writeHeader(sheet, data) {
  const base = ["Timestamp", "Delegate Number", "Delegate Status", "Election",
                "Is Test", "Is Auto-Submit", "Assisted By"];

  if (data.ballotType === "runoff") {
    sheet.appendRow([...base, "Choice"]);
  } else if (data.ballotType === "slate") {
    sheet.appendRow([...base, "Slate Vote"]);
  } else {
    // Rankings: keys are candidate letters (A, B, C…), values are ranks.
    // Sort alphabetically so column order is always A, B, C…
    const candCols = Object.keys(data.rankings || {})
      .sort()
      .map(n => "Candidate " + n + " Rank");
    sheet.appendRow([...base, ...candCols]);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
}

/**
 * Insert a new row, or overwrite the existing row for this delegate.
 * Matching is by delegate number + election key (columns B and D).
 */
function upsertRow(sheet, data) {
  const allData = sheet.getDataRange().getValues();
  let existingRow = -1;

  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][1]) === String(data.delegateNumber) &&
        String(allData[i][3]) === String(data.electionKey)) {
      existingRow = i + 1;
      break;
    }
  }

  const delegateStatusLabel =
    data.delegateStatus === "alternate" ? "Seated Alternate" :
    data.delegateStatus === "delegate"  ? "Delegate" : "";

  const base = [
    data.timestamp,
    data.delegateNumber,
    delegateStatusLabel,
    data.electionKey,
    data.isTest       ? "YES" : "NO",
    data.isAutoSubmit ? "YES" : "NO",
    data.assistedBy   || "",
  ];

  let row;
  if (data.ballotType === "runoff") {
    row = [...base, data.choice !== null && data.choice !== undefined ? data.choice : ""];
  } else if (data.ballotType === "slate") {
    row = [...base, data.slateVote || ""];
  } else {
    const sorted = Object.keys(data.rankings || {}).sort();
    const rankCols = sorted.map(n => data.rankings[n] || "");
    row = [...base, ...rankCols];
  }

  if (existingRow > 0) {
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    return true;
  } else {
    sheet.appendRow(row);
    return false;
  }
}

// ── AUDIT LOG ────────────────────────────────────────────────────────────

/**
 * Formats the vote selections into a compact human-readable string for the log.
 */
function formatVoteForLog(data) {
  if (data.ballotType === "slate") {
    return (data.slateVote || "").toUpperCase();
  }
  if (data.ballotType === "runoff") {
    return "Choice: " + (data.choice !== null && data.choice !== undefined ? data.choice : "—");
  }
  const rankings = data.rankings || {};
  return Object.entries(rankings)
    .filter(([, rank]) => rank !== null && rank !== undefined && rank !== "")
    .sort((a, b) => Number(a[1]) - Number(b[1]))
    .map(([letter, rank]) => rank + "-" + letter)
    .join(", ") || "—";
}

/**
 * Appends one row to the "Audit Log" tab for every ballot submission.
 * Creates the tab with a header if it doesn't exist yet.
 */
function appendAuditLog(ss, data, isOverwrite) {
  const LOG_TAB = "Audit Log";
  let log = ss.getSheetByName(LOG_TAB);

  if (!log) {
    log = ss.insertSheet(LOG_TAB);
    log.appendRow(["Timestamp", "Election", "Delegate ID", "Action", "Vote / Rankings"]);
    log.setFrozenRows(1);
    log.getRange(1, 1, 1, 5).setFontWeight("bold");
    log.setColumnWidth(1, 160);
    log.setColumnWidth(2,  90);
    log.setColumnWidth(3, 100);
    log.setColumnWidth(4, 160);
    log.setColumnWidth(5, 340);
  }

  log.appendRow([
    data.timestamp,
    data.electionKey,
    data.delegateNumber,
    isOverwrite ? "Re-vote (overwrote)" : "New",
    formatVoteForLog(data),
  ]);
}
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 