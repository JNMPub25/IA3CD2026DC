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

// ── UPDATE THIS ────────────────────────────────────────────────────────────
const SPREADSHEET_ID = "PASTE_YOUR_SPREADSHEET_ID_HERE";
// ──────────────────────────────────────────────────────────────────────────

/**
 * Receives POST requests from ballot.html.
 * Each election gets its own tab in the spreadsheet.
 * If a delegate re-submits, the previous row is overwritten (last submission wins).
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = JSON.parse(e.postData.contents);

    const ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = data.ballotType === "runoff"
      ? "Runoff " + data.electionKey
      : getSheetName(data.electionKey);
    let   sheet     = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      writeHeader(sheet, data);
    }

    upsertRow(sheet, data);

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

/** Also handle GET (used for connectivity testing from setup script) */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Ballot receiver is live." }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── HELPERS ────────────────────────────────────────────────────────────────

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
    // Rankings: keys are candidate numbers, values are ranks
    const candCols = Object.keys(data.rankings || {})
      .sort((a, b) => parseInt(a) - parseInt(b))
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
    // Rankings: sort by candidate number to match header order
    const sorted = Object.keys(data.rankings || {})
      .sort((a, b) => parseInt(a) - parseInt(b));
    const rankCols = sorted.map(n => data.rankings[n] || "");
    row = [...base, ...rankCols];
  }

  if (existingRow > 0) {
    // Overwrite existing submission (re-vote: last submission wins)
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
}
