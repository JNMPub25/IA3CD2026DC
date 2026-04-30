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

// ── UPDATE THIS ────────────────────────────────────────────────────────────────────
const SPREADSHEET_ID = "PASTE_YOUR_SPREADSHEET_ID_HERE";

// Name of the control tab used by the Option C timer
const CONTROL_TAB = "Control";

// Google Drive folder ID for "3rd CD Convention"
// (from the URL when you open that folder in Drive)
const CONVENTION_FOLDER_ID = "1qapYmQN_oicEyYKbuvU9zq6eP3QrNwKK";
<<<<<<< Updated upstream
// ──────────────────────────────────────────────────────────────────────────────────

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

  // moveTo replaces the deprecated addFile/removeFile pattern —
  // it moves the file into the target folder and removes it from its old parent.
  file.moveTo(folder);

  Logger.log("Moved spreadsheet " + SPREADSHEET_ID + " to 3rd CD Convention folder.");
}

// ── CONTROL TAB HELPERS (Option C timer) ──────────────────────────────────────

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

// ── GET: status polling endpoint ──────────────────────────────────────────────

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
      // Count submitted ballots: rows in the election tab minus the header row.
      // Also count any runoff tab for completeness, though usually just the main one.
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

// ── POST: ballot submissions + admin control actions ──────────────────────────
=======
// ──────────────────────────────────────────────────────────────────────────
>>>>>>> Stashed changes

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

  // moveTo replaces the deprecated addFile/removeFile pattern —
  // it moves the file into the target folder and removes it from its old parent.
  file.moveTo(folder);

  Logger.log("Moved spreadsheet " + SPREADSHEET_ID + " to 3rd CD Convention folder.");
}

// ── CONTROL TAB HELPERS (Option C timer) ──────────────────────────────────

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

// ── GET: status polling endpoint ──────────────────────────────────────────

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
      // Count submitted ballots: rows in the election tab minus the header row.
      // Also count any runoff tab for completeness, though usually just the main one.
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

// ── POST: ballot submissions + admin control actions ──────────────────────

/**
 * Receives POST requests from ballot.html.
 *
 * Two request types:
 *
 *  1. Admin control (action field present):
 *     { action: "open",  openedAt, duration, floorCount }  → sets Control tab to "open"
 *     { action: "close" }                                   → sets Control tab to "closed"
 *
 *  2. Ballot submission (no action field):
 *     Each election gets its own tab.
 *     If a delegate re-submits, the previous row is overwritten (last submission wins).
 */
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);

<<<<<<< Updated upstream
    // ── Admin: open voting ──────────────────────────────────────────────────────
=======
    // ── Admin: open voting ──────────────────────────────────────────────────
>>>>>>> Stashed changes
    if (data.action === "open") {
      const ctrl = ensureControlTab(ss);
      setControlRow(ctrl, "open", data.openedAt || "", data.duration || 0, data.floorCount || 0);
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

<<<<<<< Updated upstream
    // ── Admin: close voting ─────────────────────────────────────────────────────
=======
    // ── Admin: close voting ─────────────────────────────────────────────────
>>>>>>> Stashed changes
    if (data.action === "close") {
      const ctrl = ensureControlTab(ss);
      const row  = getControlRow(ctrl);
      setControlRow(ctrl, "closed", row.openedAt, row.duration, row.floorCount);
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

<<<<<<< Updated upstream
    // ── Ballot submission ───────────────────────────────────────────────────────
=======
    // ── Ballot submission ───────────────────────────────────────────────────
>>>>>>> Stashed changes
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

<<<<<<< Updated upstream
// ── HELPERS ────────────────────────────────────────────────────────────────────
=======
// ── HELPERS ────────────────────────────────────────────────────────────────
>>>>>>> Stashed changes

function getSheetName(electionKey) {
  const names = {
    "scc-w":   "SCC Women",
    "scc-m":   "SCC Men-NonBinary",
    "dei":     "DEI Chair",
    "scc-com": "Convention Committee",
  };
  return names[electionKey] || electionKey;
}

/**
 * Write header row based on the first submission received.
 * For ranked ballots: one column per candidate number.
 * For slate ballots:  a single "Slate Vote" column.
 */
function writeHeader(sheet, data) {
  // "Assisted By" captures the committee member's name when a ballot was
  // submitted on a delegate's behalf; blank for self-submitted ballots.
  const base = ["Timestamp", "Delegate Number", "Delegate Status", "Election", "Is Test", "Is Auto-Submit", "Assisted By"];

  if (data.ballotType === "runoff") {
    sheet.appendRow([...base, "Choice"]);
  } else if (data.ballotType === "slate") {
    sheet.appendRow([...base, "Slate Vote"]);
  } else {
    // Rankings: keys are candidate letters (A, B, C…), values are ranks.
    // Sort alphabetically so column order is always A, B, C… regardless of
    // which delegate submitted first (their shuffle order is irrelevant here).
    const candCols = Object.keys(data.rankings || {})
      .sort()
      .map(n => `Candidate ${n} Rank`);
    sheet.appendRow([...base, ...candCols]);
  }

  // Freeze header row and bold it
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight("bold");
}

/**
 * Insert a new row, or overwrite the existing row for this delegate.
 * Matching is by delegate number + election key (columns B and C).
 */
function upsertRow(sheet, data) {
  const allData = sheet.getDataRange().getValues();
  let existingRow = -1;

  // Skip header (row 0 in 0-indexed), look for matching delegate.
  // Column layout: A=Timestamp, B=Delegate Number, C=Delegate Status,
  //                D=Election, E=Is Test, F=Is Auto-Submit, G=Assisted By
  // → index 1 = Delegate Number, index 3 = Election
  for (let i = 1; i < allData.length; i++) {
    if (String(allData[i][1]) === String(data.delegateNumber) &&
        String(allData[i][3]) === String(data.electionKey)) {
      existingRow = i + 1;  // 1-indexed for Sheets API
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
    data.assistedBy   || "",   // blank for self-submitted ballots
  ];

  let row;
  if (data.ballotType === "runoff") {
    row = [...base, data.choice !== null && data.choice !== undefined ? data.choice : ""];
  } else if (data.ballotType === "slate") {
    row = [...base, data.slateVote || ""];
  } else {
    // Rankings: sort alphabetically (A, B, C…) to match header column order
    const sorted = Object.keys(data.rankings || {})
      .sort();
    const rankCols = sorted.map(n => data.rankings[n] || "");
    row = [...base, ...rankCols];
  }

  if (existingRow > 0) {
    // Overwrite existing submission (re-vote: last submission wins)
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    return true;   // was a re-vote (overwrite)
  } else {
    sheet.appendRow(row);
    return false;  // was a new submission
  }
}

<<<<<<< Updated upstream
// ── AUDIT LOG ──────────────────────────────────────────────────────────────────
=======
// ── AUDIT LOG ──────────────────────────────────────────────────────────────
>>>>>>> Stashed changes

/**
 * Formats the vote selections into a compact human-readable string for the log.
 *
 * Ranked:  "1-B, 2-F, 3-D, 4-A, 5-C, 6-E"   (sorted by rank ascending)
 * Slate:   "YES" | "NO"
 * Runoff:  "Choice: B"
 */
function formatVoteForLog(data) {
  if (data.ballotType === "slate") {
    return (data.slateVote || "").toUpperCase();
  }
  if (data.ballotType === "runoff") {
    return "Choice: " + (data.choice !== null && data.choice !== undefined ? data.choice : "—");
  }
  // Ranked: build "rank-letter" pairs, sort by rank number ascending
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
 *
 * Columns: Timestamp | Election | Delegate ID | Action | Vote / Rankings
 *
 * Action is "New" for first-time submissions and "Re-vote (overwrote)" when
 * a delegate's earlier ballot was replaced.
 */
function appendAuditLog(ss, data, isOverwrite) {
  const LOG_TAB = "Audit Log";
  let log = ss.getSheetByName(LOG_TAB);

  if (!log) {
    log = ss.insertSheet(LOG_TAB);
    log.appendRow([
      "Timestamp", "Election", "Delegate ID", "Action", "Vote / Rankings"
    ]);
    log.setFrozenRows(1);
    log.getRange(1, 1, 1, 5).setFontWeight("bold");
    // Widen columns for readability
    log.setColumnWidth(1, 160);  // Timestamp
    log.setColumnWidth(2,  90);  // Election
    log.setColumnWidth(3, 100);  // Delegate ID
    log.setColumnWidth(4, 160);  // Action
    log.setColumnWidth(5, 340);  // Vote / Rankings
  }

  log.appendRow([
    data.timestamp,
    data.electionKey,
    data.delegateNumber,
    isOverwrite ? "Re-vote (overwrote)" : "New",
    formatVoteForLog(data),
  ]);
}
