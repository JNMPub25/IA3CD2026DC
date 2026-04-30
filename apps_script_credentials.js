// ════════════════════════════════════════════════════════════════════════
//  3rd CD Convention 2026 — Credentials Tracker Apps Script Backend
//  Deploy as: Execute as Me | Who has access: Anyone
//  Separate deployment from the ballot Apps Script
//
//  Google Sheet tabs required:
//    1. "Roster"                  — delegates & alternates (upload from CSV)
//    2. "Master Credentials List" — all entries from CCM terminals
//    3. "Quorum"                  — live quorum dashboard
//    4. "Reassignments Dashboard"  — county seating status, priority matching, reassignment pool
//    5. "Credentials Audit Log"   — all changes (terminal syncs, edits, Chair edits)
// ════════════════════════════════════════════════════════════════════════

var SS_ID = ''; // ← PASTE YOUR GOOGLE SHEET ID HERE after creating the sheet

// Sheet name constants
var SHEET_ROSTER     = 'Roster';
var SHEET_MASTER     = 'Master Credentials List';
var SHEET_QUORUM     = 'Quorum';
var SHEET_REASSIGN   = 'Reassignments Dashboard';

// Urban vs Rural classification (Convention Call Book priority tiers c/d)
var URBAN_COUNTIES = ['Polk', 'Dallas'];
var SHEET_AUDIT      = 'Credentials Audit Log';

// Column layout for Master Credentials List
var MCL_COLS = {
  entryId:    1,   // A — UUID from terminal
  timestamp:  2,   // B — CT datetime string
  ccmName:    3,   // C — Credentials Committee Member name
  ccmId:      4,   // D — Terminal identifier
  reason:     5,   // E — Reason code
  fromId:     6,   // F — From ID (Badge ID leaving/yielding)
  toId:       7,   // G — To / Arriving ID
  notes:      8,   // H — Notes
  isEdit:     9,   // I — TRUE if this was an edit to a prior entry
  origValues: 10,  // J — JSON of original values (for edits)
  syncedAt:   11   // K — Timestamp when record arrived at server
};

// Column layout for Audit Log
var AUDIT_COLS = {
  auditTime:   1,  // A — When audit record created (CT)
  actionType:  2,  // B — NEW | EDIT-TERMINAL | EDIT-CHAIR | DELETE-CHAIR
  entryId:     3,  // C — UUID of the affected entry
  ccmOrChair:  4,  // D — Who made the change
  changedFields: 5, // E — JSON of what changed (old → new)
  reason:      6,  // F — Reason code at time of action
  fromId:      7,  // G
  toId:        8,  // H
  notes:       9   // I
};

// ════════════════════════════════════════════════════════════════════════
//  GET HANDLER
// ════════════════════════════════════════════════════════════════════════
function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : '';
  var result;

  try {
    if (action === 'getRoster') {
      result = getRosterData();
    } else if (action === 'checkCredentials') {
      var delegateId = e.parameter.delegateId || '';
      result = checkDelegateReassigned(delegateId);
    } else if (action === 'getQuorum') {
      result = getQuorumData();
    } else {
      result = { error: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { error: err.toString() };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════════
//  POST HANDLER — receives sync payloads from CCM terminals
// ════════════════════════════════════════════════════════════════════════
function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;

    if (action === 'syncEntries') {
      syncEntries(payload);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ════════════════════════════════════════════════════════════════════════
//  SYNC ENTRIES (POST action)
// ════════════════════════════════════════════════════════════════════════
function syncEntries(payload) {
  var ss      = SpreadsheetApp.openById(SS_ID);
  var mclSheet   = getOrCreateSheet(ss, SHEET_MASTER,  getMCLHeaders());
  var auditSheet = getOrCreateSheet(ss, SHEET_AUDIT,   getAuditHeaders());

  var entries  = payload.entries  || [];
  var ccmName  = payload.ccmName  || 'Unknown CCM';
  var ccmId    = payload.ccmId    || '';
  var syncedAt = getCTTimestamp();

  // Build a map of existing entry IDs for fast lookup
  var existingRows = mclSheet.getDataRange().getValues();
  var idToRow = {};
  var dedupSet = {};  // reason|fromId|toId → true (for detecting duplicate Badge ID + Reason)
  for (var i = 1; i < existingRows.length; i++) {
    var rowId = existingRows[i][MCL_COLS.entryId - 1];
    if (rowId) idToRow[rowId] = i + 1; // 1-based row number
    // Build dedup key from existing MCL rows
    var existReason = String(existingRows[i][MCL_COLS.reason - 1] || '').trim();
    var existFrom   = String(existingRows[i][MCL_COLS.fromId - 1] || '').trim();
    var existTo     = String(existingRows[i][MCL_COLS.toId   - 1] || '').trim();
    if (existReason) dedupSet[existReason + '|' + existFrom + '|' + existTo] = true;
  }

  for (var j = 0; j < entries.length; j++) {
    var entry = entries[j];
    var entryId = entry.entryId || '';
    var isEdit  = !!entry.isEdit;
    var origRow = idToRow[entryId];

    if (isEdit && origRow) {
      // UPDATE existing row — batch into one setValues call for performance
      var origVals = entry.originalValues ? JSON.stringify(entry.originalValues) : '';
      mclSheet.getRange(origRow, MCL_COLS.reason, 1, 6).setValues([[
        entry.reason || '',
        entry.fromId || '',
        entry.toId   || '',
        entry.notes  || '',
        true,
        origVals
      ]]);

      // Audit: EDIT-TERMINAL
      appendAuditRow(auditSheet, {
        auditTime:     syncedAt,
        actionType:    'EDIT-TERMINAL',
        entryId:       entryId,
        ccmOrChair:    ccmName + ' [' + ccmId + ']',
        changedFields: buildChangedFields(entry.originalValues, entry),
        reason:        entry.reason || '',
        fromId:        entry.fromId || '',
        toId:          entry.toId   || '',
        notes:         entry.notes  || ''
      });
    } else {
      // ── Dedup check: same Badge ID + same Reason already in MCL? ──
      var dedupKey = (entry.reason || '').trim() + '|' + (entry.fromId || '').trim() + '|' + (entry.toId || '').trim();
      if (dedupSet[dedupKey]) {
        // Duplicate — log to Audit only, do NOT add to MCL
        appendAuditRow(auditSheet, {
          auditTime:     syncedAt,
          actionType:    'DUPLICATE',
          entryId:       entryId,
          ccmOrChair:    ccmName + ' [' + ccmId + ']',
          changedFields: 'Duplicate of existing entry: ' + dedupKey,
          reason:        entry.reason  || '',
          fromId:        entry.fromId  || '',
          toId:          entry.toId    || '',
          notes:         entry.notes   || ''
        });
      } else {
        // NEW row (also catches isEdit=true when the row isn't found — treat as NEW)
        var newRow = [
          entryId,
          entry.datetime || syncedAt,
          ccmName,
          ccmId,
          entry.reason  || '',
          entry.fromId  || '',
          entry.toId    || '',
          entry.notes   || '',
          false,
          '',
          syncedAt
        ];
        mclSheet.appendRow(newRow);

        // Track this new row in our map and dedup set
        var newRowNum = mclSheet.getLastRow();
        idToRow[entryId] = newRowNum;
        dedupSet[dedupKey] = true;

        // Audit: NEW
        appendAuditRow(auditSheet, {
          auditTime:     syncedAt,
          actionType:    'NEW',
          entryId:       entryId,
          ccmOrChair:    ccmName + ' [' + ccmId + ']',
          changedFields: '',
          reason:        entry.reason  || '',
          fromId:        entry.fromId  || '',
          toId:          entry.toId    || '',
          notes:         entry.notes   || ''
        });
      }
    }
    // (isEdit=true but row not found is handled above as a new insert)
  }

  // Refresh computed sheets after sync
  refreshQuorumSheet(ss);
  refreshReassignmentSheet(ss);
}

// ════════════════════════════════════════════════════════════════════════
//  ROSTER — fetch from Roster sheet
// ════════════════════════════════════════════════════════════════════════
function getRosterData() {
  var ss    = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName(SHEET_ROSTER);
  if (!sheet) return { roster: [], error: 'Roster sheet not found' };

  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { roster: [] };

  // Detect columns by header name (case-insensitive)
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var idIdx      = findCol(headers, ['badge id', 'id', 'badge_id']);
  var lastIdx    = findCol(headers, ['last name', 'last', 'lastname']);
  var firstIdx   = findCol(headers, ['first name', 'first', 'firstname']);
  var countyIdx  = findCol(headers, ['county']);

  if (idIdx < 0 || lastIdx < 0 || firstIdx < 0) return { roster: [], error: 'Required columns not found' };

  var roster = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var badgeId = String(row[idIdx] || '').trim();
    if (!badgeId) continue;
    roster.push({
      id:     badgeId,
      last:   String(row[lastIdx]  || '').trim(),
      first:  String(row[firstIdx] || '').trim(),
      county: countyIdx >= 0 ? String(row[countyIdx] || '').trim() : ''
    });
  }
  return { roster: roster };
}

function findCol(headers, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = headers.indexOf(names[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}

// ════════════════════════════════════════════════════════════════════════
//  CREDENTIALS REASSIGNED CHECK
// ════════════════════════════════════════════════════════════════════════
function checkDelegateReassigned(delegateId) {
  if (!delegateId) return { reassigned: false };
  var ss    = SpreadsheetApp.openById(SS_ID);
  var sheet = ss.getSheetByName(SHEET_MASTER);
  if (!sheet) return { reassigned: false };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var reason = String(data[i][MCL_COLS.reason - 1] || '');
    var fromId = String(data[i][MCL_COLS.fromId - 1] || '');
    if (reason === 'Seating-Regular Alternates' && fromId === delegateId) {
      return { reassigned: true, delegateId: delegateId };
    }
  }
  return { reassigned: false, delegateId: delegateId };
}

// ════════════════════════════════════════════════════════════════════════
//  QUORUM DATA
// ════════════════════════════════════════════════════════════════════════
function getQuorumData() {
  var ss = SpreadsheetApp.openById(SS_ID);
  return computeQuorum(ss);
}

function computeQuorum(ss) {
  var sheet = ss.getSheetByName(SHEET_MASTER);
  if (!sheet) return { seatedDelegates: 0, totalSeated: 0, quorumAchieved: false };

  var data = sheet.getDataRange().getValues();
  // Track unique D-#### IDs currently active
  var activeD = {};   // delegateId → true if active on floor

  for (var i = 1; i < data.length; i++) {
    var reason = String(data[i][MCL_COLS.reason - 1] || '');
    var fromId = String(data[i][MCL_COLS.fromId - 1] || '').trim();
    var toId   = String(data[i][MCL_COLS.toId   - 1] || '').trim();

    if (reason === 'Arriving' && toId.startsWith('D-')) {
      activeD[toId] = true;
    } else if ((reason === 'Seating-Preferred Alternate' || reason === 'Seating-Regular Alternates') && fromId.startsWith('D-')) {
      // A delegate seat is filled (by alternate on their credentials)
      activeD[fromId] = true;
    } else if (reason === 'Surrender' && fromId.startsWith('D-')) {
      delete activeD[fromId];
    } else if (reason === 'Reclaiming' && toId.startsWith('D-')) {
      // Delegate reclaims — seat active again under their ID
      activeD[toId] = true;
    }
  }

  var seatedCount = Object.keys(activeD).length;
  var quorumMet   = seatedCount >= 98;

  return {
    seatedDelegates: seatedCount,
    totalSeated:     seatedCount, // can be expanded to include all seated
    quorumNeeded:    98,
    electedDelegates: 243,
    quorumAchieved:  quorumMet
  };
}

// ════════════════════════════════════════════════════════════════════════
//  QUORUM SHEET — refresh formatted dashboard
// ════════════════════════════════════════════════════════════════════════
function refreshQuorumSheet(ss) {
  var sheet  = getOrCreateSheet(ss, SHEET_QUORUM, null);
  var q      = computeQuorum(ss);
  var mcl    = ss.getSheetByName(SHEET_MASTER);
  var data   = mcl ? mcl.getDataRange().getValues() : [[]];

  // Count seated delegates — mirrors computeQuorum logic, including surrenders
  var activeD = {};   // track unique D-#### IDs currently on the floor
  var seatedDelegates = 0, seatedPref = 0, seatedReg = 0;
  for (var i = 1; i < data.length; i++) {
    var reason = String(data[i][MCL_COLS.reason - 1] || '');
    var fromId = String(data[i][MCL_COLS.fromId - 1] || '').trim();
    var toId   = String(data[i][MCL_COLS.toId   - 1] || '').trim();
    if (reason === 'Arriving' && toId.startsWith('D-'))                       { activeD[toId] = true; seatedDelegates++; }
    if (reason === 'Seating-Preferred Alternate' && fromId.startsWith('D-'))  { activeD[fromId] = true; seatedPref++; }
    if (reason === 'Seating-Regular Alternates'  && fromId.startsWith('D-'))  { activeD[fromId] = true; seatedReg++; }
    if (reason === 'Surrender' && fromId.startsWith('D-'))                    { delete activeD[fromId]; }
    if (reason === 'Reclaiming' && toId.startsWith('D-'))                     { activeD[toId] = true; }
  }
  var totalSeated = Object.keys(activeD).length;
  var quorumPct   = (totalSeated / 243 * 100).toFixed(1) + '%';
  var achieved    = totalSeated >= 98;
  var now         = getCTTimestamp();

  sheet.clearContents();
  sheet.clearFormats();

  // Row 1: Title
  sheet.getRange('A1').setValue('3rd CD Convention 2026 — Credentials Quorum Dashboard').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue('Last updated: ' + now).setFontColor('#666666').setFontSize(10);
  sheet.getRange('A3').setValue('');

  // Row 4–5: Quorum summary
  sheet.getRange('A4').setValue('Quorum:');
  sheet.getRange('B4').setValue(quorumPct + '  (' + totalSeated + ' / 243 elected delegates)');
  if (achieved) {
    sheet.getRange('B4').setBackground('#c6efce').setFontColor('#276221').setFontWeight('bold');
    sheet.getRange('C4').setValue('✓ QUORUM ACHIEVED — ' + now).setFontWeight('bold').setFontColor('#276221');
  } else {
    sheet.getRange('B4').setBackground('#ffc7ce').setFontColor('#9c0006').setFontWeight('bold');
  }

  sheet.getRange('A5').setValue('Needs for Quorum (40% of 243):').setFontWeight('bold');
  sheet.getRange('B5').setValue(98);
  sheet.getRange('A6').setValue('');

  // Row 7+: Detail counts
  var rows = [
    ['Elected Delegates (from Roster):', 243],
    ['Seated Delegates (Arriving, D-####):', seatedDelegates],
    ['Seated Preferred Alternates (Seating-Preferred Alt, From=D-####):', seatedPref],
    ['Seated Regular Alternates (Seating-Regular Alt, From=D-####):', seatedReg],
    ['Total Seated:', totalSeated],
  ];
  for (var r = 0; r < rows.length; r++) {
    sheet.getRange(7 + r, 1).setValue(rows[r][0]).setFontWeight(r === rows.length - 1 ? 'bold' : 'normal');
    sheet.getRange(7 + r, 2).setValue(rows[r][1]).setFontWeight(r === rows.length - 1 ? 'bold' : 'normal');
  }

  // Row 13+: Reassignment counts (from Reassignment Options sheet)
  sheet.getRange('A13').setValue('');
  sheet.getRange('A14').setValue('Unclaimed Delegate IDs (from Reassignments Dashboard):').setFontWeight('bold');
  sheet.getRange('B14').setFormula("=COUNTA('Reassignments Dashboard'!A:A)-1");
  sheet.getRange('A15').setValue('Available Alternates (from Reassignments Dashboard):').setFontWeight('bold');
  sheet.getRange('B15').setFormula("=COUNTA('Reassignments Dashboard'!E:E)-1");

  // Auto-resize
  sheet.autoResizeColumns(1, 4);
}

// ════════════════════════════════════════════════════════════════════════
//  REASSIGNMENT OPTIONS SHEET
// ════════════════════════════════════════════════════════════════════════
function refreshReassignmentSheet(ss) {
  var reassSheet = getOrCreateSheet(ss, SHEET_REASSIGN, null);
  var rosterSheet = ss.getSheetByName(SHEET_ROSTER);
  var mclSheet    = ss.getSheetByName(SHEET_MASTER);

  if (!rosterSheet || !mclSheet) return;

  var rosterData = rosterSheet.getDataRange().getValues();
  var mclData    = mclSheet.getDataRange().getValues();

  var headers   = rosterData[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var idIdx     = findCol(headers, ['badge id', 'id', 'badge_id']);
  var lastIdx   = findCol(headers, ['last name', 'last', 'lastname']);
  var firstIdx  = findCol(headers, ['first name', 'first', 'firstname']);
  var countyIdx = findCol(headers, ['county']);
  var genderIdx = findCol(headers, ['gender']);

  if (idIdx < 0) return;

  // ── Build sets from MCL (mirrors computeQuorum logic, including Surrender/Reclaiming) ──
  var seatedDelegateIds   = {};  // D-#### whose seat is currently active
  var arrivedAlternateIds = {};  // A-#### who have physically arrived
  var seatedAlternateIds  = {};  // A-#### who have been assigned to a delegate seat

  for (var i = 1; i < mclData.length; i++) {
    var reason = String(mclData[i][MCL_COLS.reason - 1] || '');
    var fromId = String(mclData[i][MCL_COLS.fromId - 1] || '').trim();
    var toId   = String(mclData[i][MCL_COLS.toId   - 1] || '').trim();

    // Delegate seat tracking (with Surrender/Reclaiming)
    if (reason === 'Arriving' && toId.startsWith('D-'))                      seatedDelegateIds[toId] = true;
    if (reason === 'Seating-Preferred Alternate' && fromId.startsWith('D-')) seatedDelegateIds[fromId] = true;
    if (reason === 'Seating-Regular Alternates'  && fromId.startsWith('D-')) seatedDelegateIds[fromId] = true;
    if (reason === 'Surrender' && fromId.startsWith('D-'))                   delete seatedDelegateIds[fromId];
    if (reason === 'Reclaiming' && toId.startsWith('D-'))                    seatedDelegateIds[toId] = true;

    // Alternate tracking
    if (reason === 'Arriving' && toId.startsWith('A-')) arrivedAlternateIds[toId] = true;
    if (reason === 'Seating-Preferred Alternate' && toId.startsWith('A-')) seatedAlternateIds[toId] = true;
    if (reason === 'Seating-Regular Alternates'  && toId.startsWith('A-')) seatedAlternateIds[toId] = true;
  }

  // ── Build roster lookups ──
  var rosterById = {};  // badgeId → {id, county, gender, name, isUrban}
  var unassignedDelegates = [];
  var availableAlternates = [];
  var countyDelegateTotal = {};   // county → total elected delegate count
  var countyDelegateSeated = {};  // county → seated delegate count

  for (var j = 1; j < rosterData.length; j++) {
    var row     = rosterData[j];
    var badgeId = String(row[idIdx] || '').trim();
    if (!badgeId) continue;
    var county  = countyIdx >= 0 ? String(row[countyIdx] || '').trim() : '';
    var gender  = genderIdx >= 0 ? String(row[genderIdx] || '').trim() : '';
    var name    = String(row[lastIdx] || '').trim() + ', ' + String(row[firstIdx] || '').trim();
    var isUrban = URBAN_COUNTIES.indexOf(county) >= 0;

    rosterById[badgeId] = { id: badgeId, county: county, gender: gender, name: name, isUrban: isUrban };

    if (badgeId.startsWith('D-')) {
      // County delegate counts
      if (!countyDelegateTotal[county])  countyDelegateTotal[county]  = 0;
      if (!countyDelegateSeated[county]) countyDelegateSeated[county] = 0;
      countyDelegateTotal[county]++;
      if (seatedDelegateIds[badgeId]) {
        countyDelegateSeated[county]++;
      } else {
        unassignedDelegates.push({ id: badgeId, county: county, gender: gender, name: name, isUrban: isUrban });
      }
    } else if (badgeId.startsWith('A-') && arrivedAlternateIds[badgeId] && !seatedAlternateIds[badgeId]) {
      availableAlternates.push({ id: badgeId, county: county, gender: gender, name: name, isUrban: isUrban });
    }
  }

  // ── Clear sheet ──
  reassSheet.clearContents();
  reassSheet.clearFormats();

  var now     = getCTTimestamp();
  var curRow  = 1;
  var NAVY    = '#003366';
  var WHITE   = '#ffffff';
  var GREEN   = '#c6efce';
  var LTBLUE  = '#dce6f1';

  // ════════════════════════════════════════════════════════════════
  //  SECTION 1 — County Seating Status (Visual Box Grid)
  // ════════════════════════════════════════════════════════════════
  reassSheet.getRange(curRow, 1).setValue('COUNTY SEATING STATUS').setFontWeight('bold').setFontSize(12).setFontColor(NAVY);
  curRow++;
  reassSheet.getRange(curRow, 1).setValue('Last updated: ' + now).setFontColor('#666666').setFontSize(9);
  curRow++;

  // Summary totals
  var totalElected = 0; var totalSeatedAll = 0;
  for (var ck in countyDelegateTotal)  totalElected   += countyDelegateTotal[ck];
  for (var cs in countyDelegateSeated) totalSeatedAll  += countyDelegateSeated[cs];
  var totalPct = Math.round((totalSeatedAll / (totalElected || 1)) * 100);
  reassSheet.getRange(curRow, 1).setValue('Total: ' + totalSeatedAll + ' / ' + totalElected + ' seated (' + totalPct + '%)').setFontWeight('bold').setFontSize(10);
  curRow++;

  // Color helper: returns [bgColor, textColor] based on seating %
  function boxColors(pct) {
    if (pct === 100) return ['#27500A', '#EAF3DE'];  // dark green, light text
    if (pct >= 50)  return ['#639922', '#EAF3DE'];   // med green, light text
    if (pct > 0)    return ['#C0DD97', '#173404'];   // light green, dark text
    return ['#F0997B', '#4A1B0C'];                    // coral, dark text
  }

  // Geographic grid layout — [countyName, startCol, colSpan (default 2)]
  // Each box = 2 rows tall. Cols are 1-indexed. Max 14 cols wide.
  var gridRows = [
    // Tier 1 (North): Guthrie, Greene, Dallas, Polk (Polk double-wide)
    [['Guthrie',1], ['Greene',3], ['Dallas',5], ['Polk',7,4]],
    // Tier 2 (Central-North): Cass, Adair, Madison
    [['Cass',1], ['Adair',3], ['Madison',5]],
    // Tier 3 (Central): Montgomery thru Wapello
    [['Montgomery',1], ['Adams',3], ['Union',5], ['Clarke',7], ['Lucas',9], ['Monroe',11], ['Wapello',13]],
    // Tier 4 (South border): Page thru Davis
    [['Page',1], ['Taylor',3], ['Ringgold',5], ['Decatur',7], ['Wayne',9], ['Appanoose',11], ['Davis',13]]
  ];

  var gridStartRow = curRow;
  for (var gr = 0; gr < gridRows.length; gr++) {
    var tier = gridRows[gr];
    var rowTop = gridStartRow + (gr * 2); // each tier = 2 sheet rows

    for (var gc = 0; gc < tier.length; gc++) {
      var ctyName  = tier[gc][0];
      var startCol = tier[gc][1];
      var colSpan  = tier[gc][2] || 2;
      var total    = countyDelegateTotal[ctyName]  || 0;
      var seated   = countyDelegateSeated[ctyName] || 0;
      var pct      = total > 0 ? Math.round((seated / total) * 100) : 0;
      var colors   = boxColors(pct);

      // Merge the box cells (2 rows x colSpan cols)
      var boxRange = reassSheet.getRange(rowTop, startCol, 2, colSpan);
      boxRange.merge()
        .setValue(ctyName + '\n' + seated + '/' + total + ' (' + pct + '%)')
        .setBackground(colors[0])
        .setFontColor(colors[1])
        .setFontWeight('bold')
        .setFontSize(9)
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
        .setBorder(true, true, true, true, false, false, WHITE, SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  // Legend row after the grid
  curRow = gridStartRow + (gridRows.length * 2) + 1;
  var legendColors = [['#27500A','100%'], ['#639922','50-99%'], ['#C0DD97','1-49%'], ['#F0997B','0%']];
  for (var lg = 0; lg < legendColors.length; lg++) {
    var lCol = 1 + (lg * 3);
    reassSheet.getRange(curRow, lCol).setBackground(legendColors[lg][0]).setValue('  ').setFontSize(8);
    reassSheet.getRange(curRow, lCol + 1).setValue(legendColors[lg][1]).setFontSize(8).setFontColor('#666666');
  }
  curRow += 2; // blank row after legend

  // ════════════════════════════════════════════════════════════════
  //  SECTION 2 — Priority-Matched Suggestions (Convention Call Book I-D.4)
  // ════════════════════════════════════════════════════════════════
  reassSheet.getRange(curRow, 1).setValue('PRIORITY-MATCHED SUGGESTIONS').setFontWeight('bold').setFontSize(12).setFontColor(NAVY);
  curRow++;
  reassSheet.getRange(curRow, 1).setValue('Convention Call Book I-D.4 — 6-tier priority order. Top 3 alternates per open seat.').setFontColor('#666666').setFontSize(9);
  curRow++;

  if (unassignedDelegates.length === 0) {
    reassSheet.getRange(curRow, 1).setValue('No open delegate seats.').setFontStyle('italic').setFontColor('#666666');
    curRow += 2;
  } else if (availableAlternates.length === 0) {
    reassSheet.getRange(curRow, 1).setValue('No available alternates to match.').setFontStyle('italic').setFontColor('#666666');
    curRow += 2;
  } else {
    var matchHeaders = [
      'Open Seat ID', 'Delegate Name', 'Seat County', 'Seat Gender', 'Rural/Urban',
      '#1 Match ID', '#1 Name', '#1 County', '#1 Gender', '#1 Priority Tier',
      '#2 Match ID', '#2 Name', '#2 County', '#2 Gender', '#2 Priority Tier',
      '#3 Match ID', '#3 Name', '#3 County', '#3 Gender', '#3 Priority Tier'
    ];
    reassSheet.getRange(curRow, 1, 1, matchHeaders.length).setValues([matchHeaders])
      .setFontWeight('bold').setBackground(NAVY).setFontColor(WHITE);
    curRow++;

    // Pre-score all open seats so we can sort by best-match tier before displaying
    var seatScores = [];
    for (var d = 0; d < unassignedDelegates.length; d++) {
      var seat = unassignedDelegates[d];
      var scored = [];

      for (var a = 0; a < availableAlternates.length; a++) {
        var alt  = availableAlternates[a];
        var tier = computePriorityTier(seat, alt);
        scored.push({ alt: alt, tier: tier });
      }
      scored.sort(function(x, y) { return x.tier - y.tier; });
      seatScores.push({ seat: seat, scored: scored, bestTier: scored.length > 0 ? scored[0].tier : 99 });
    }

    // Sort seats so the best (lowest-tier) matches appear first
    seatScores.sort(function(x, y) { return x.bestTier - y.bestTier; });

    // Display at most 10 rows
    var displayLimit = Math.min(10, seatScores.length);
    var tierLabels = ['1: Same gender, same county', '2: Diff gender, same county',
      '3: Same gender, similar rural/urban', '4: Diff gender, similar rural/urban',
      '5: Same gender, any county', '6: Diff gender, any county'];

    for (var d = 0; d < displayLimit; d++) {
      var entry = seatScores[d];
      var rowData = [entry.seat.id, entry.seat.name, entry.seat.county, entry.seat.gender, entry.seat.isUrban ? 'Urban' : 'Rural'];

      for (var m = 0; m < 3; m++) {
        if (m < entry.scored.length) {
          var match = entry.scored[m];
          rowData.push(match.alt.id, match.alt.name, match.alt.county, match.alt.gender, tierLabels[match.tier - 1] || ('Tier ' + match.tier));
        } else {
          rowData.push('', '', '', '', '');
        }
      }

      reassSheet.getRange(curRow, 1, 1, rowData.length).setValues([rowData]);

      // Color-code tier 1 matches
      if (entry.scored.length > 0 && entry.scored[0].tier === 1) {
        reassSheet.getRange(curRow, 6, 1, 5).setBackground(GREEN);
      }
      curRow++;
    }

    // Overflow note if more open seats exist beyond the displayed 10
    if (seatScores.length > displayLimit) {
      reassSheet.getRange(curRow, 1).setValue('+ ' + (seatScores.length - displayLimit) + ' more open seats — see full list in Reassignment Pool below.')
        .setFontStyle('italic').setFontColor('#666666').setFontSize(9);
      curRow++;
    }
    curRow++; // blank row
  }

  // ════════════════════════════════════════════════════════════════
  //  SECTION 3 — Reassignment Pool (replaces old Reassignment Options content)
  // ════════════════════════════════════════════════════════════════
  reassSheet.getRange(curRow, 1).setValue('REASSIGNMENT POOL').setFontWeight('bold').setFontSize(12).setFontColor(NAVY);
  curRow++;

  var poolHeaders = ['Unassigned Delegate ID', 'Delegate County', 'Delegate Gender', 'Delegate Name',
    '', 'Available Alternate ID', 'Alternate County', 'Alternate Gender', 'Alternate Name'];
  reassSheet.getRange(curRow, 1, 1, poolHeaders.length).setValues([poolHeaders])
    .setFontWeight('bold').setBackground(NAVY).setFontColor(WHITE);
  curRow++;

  var maxRows = Math.max(unassignedDelegates.length, availableAlternates.length);
  if (maxRows > 0) {
    var batchData = [];
    for (var r = 0; r < maxRows; r++) {
      var dObj = unassignedDelegates[r];
      var aObj = availableAlternates[r];
      batchData.push([
        dObj ? dObj.id     : '', dObj ? dObj.county : '', dObj ? dObj.gender : '', dObj ? dObj.name : '',
        '',
        aObj ? aObj.id     : '', aObj ? aObj.county : '', aObj ? aObj.gender : '', aObj ? aObj.name : ''
      ]);
    }
    reassSheet.getRange(curRow, 1, batchData.length, 9).setValues(batchData);
    curRow += batchData.length;
  } else {
    reassSheet.getRange(curRow, 1).setValue('No unassigned delegates or available alternates.').setFontStyle('italic').setFontColor('#666666');
    curRow++;
  }

  // Set compact column widths (avoid overly wide columns from long tier labels)
  var colWidths = [
    90,   // A: Seat ID / County name
    120,  // B: Delegate Name / name
    85,   // C: County
    65,   // D: Gender
    70,   // E: Rural/Urban
    80,   // F: #1 Match ID
    120,  // G: #1 Name
    80,   // H: #1 County
    55,   // I: #1 Gender
    170,  // J: #1 Priority Tier
    80,   // K: #2 Match ID
    120,  // L: #2 Name
    80,   // M: #2 County
    55,   // N: #2 Gender
    170,  // O: #2 Priority Tier
    80,   // P: #3 Match ID
    120,  // Q: #3 Name
    80,   // R: #3 County
    55,   // S: #3 Gender
    170   // T: #3 Priority Tier
  ];
  for (var w = 0; w < colWidths.length; w++) {
    reassSheet.setColumnWidth(w + 1, colWidths[w]);
  }
}

// ── Priority tier computation (Convention Call Book I-D.4) ──
function computePriorityTier(seat, alt) {
  var sameGender  = seat.gender !== '' && alt.gender !== '' && seat.gender === alt.gender;
  var sameCounty  = seat.county !== '' && alt.county !== '' && seat.county === alt.county;
  var sameRuralUrban = seat.isUrban === alt.isUrban;

  if (sameCounty && sameGender)      return 1;
  if (sameCounty && !sameGender)     return 2;
  if (sameRuralUrban && sameGender)  return 3;
  if (sameRuralUrban && !sameGender) return 4;
  if (sameGender)                    return 5;
  return 6;
}

// ════════════════════════════════════════════════════════════════════════
//  CHAIR EDIT DETECTION (installable trigger)
//  Install this as an installable trigger on the spreadsheet:
//  Extensions → Apps Script → Triggers → Add trigger → onEditTrigger → spreadsheet → onEdit
//  NOTE: Named onEditTrigger (not onEdit) to avoid firing as both a simple
//        trigger AND an installable trigger, which would create duplicate audit rows.
// ════════════════════════════════════════════════════════════════════════
function onEditTrigger(e) {
  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_MASTER) return;

  var row = e.range.getRow();
  if (row <= 1) return; // header row

  var ss         = SpreadsheetApp.openById(SS_ID);
  var auditSheet = getOrCreateSheet(ss, SHEET_AUDIT, getAuditHeaders());
  var entryId    = sheet.getRange(row, MCL_COLS.entryId).getValue() || ('DIRECT-ROW-' + row);

  var oldVal = e.oldValue !== undefined ? e.oldValue : '(unknown)';
  var newVal = e.value    !== undefined ? e.value    : e.range.getValue();
  var colIdx = e.range.getColumn();
  var colName = getColName(colIdx);

  appendAuditRow(auditSheet, {
    auditTime:     getCTTimestamp(),
    actionType:    'EDIT-CHAIR',
    entryId:       entryId,
    ccmOrChair:    Session.getActiveUser().getEmail() || 'Chair/Designee',
    changedFields: colName + ': "' + oldVal + '" → "' + newVal + '"',
    reason:        sheet.getRange(row, MCL_COLS.reason).getValue(),
    fromId:        sheet.getRange(row, MCL_COLS.fromId).getValue(),
    toId:          sheet.getRange(row, MCL_COLS.toId).getValue(),
    notes:         sheet.getRange(row, MCL_COLS.notes).getValue()
  });
}

function getColName(colIdx) {
  var names = {1:'Entry ID', 2:'Timestamp', 3:'CCM Name', 4:'CCM ID', 5:'Reason', 6:'From ID', 7:'To ID', 8:'Notes', 9:'Is Edit', 10:'Original Values', 11:'Synced At'};
  return names[colIdx] || ('Column ' + colIdx);
}

// ════════════════════════════════════════════════════════════════════════
//  HELPERS
// ════════════════════════════════════════════════════════════════════════
function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#003366').setFontColor('#ffffff');
    }
  }
  return sheet;
}

function getMCLHeaders() {
  return ['Entry ID', 'Timestamp (CT)', 'CCM Name', 'CCM Terminal ID', 'Reason', 'From ID', 'To / Arriving ID', 'Notes', 'Is Edit', 'Original Values (JSON)', 'Synced At (CT)'];
}

function getAuditHeaders() {
  return ['Audit Timestamp (CT)', 'Action Type', 'Entry ID', 'By (CCM / Chair)', 'Changed Fields', 'Reason', 'From ID', 'To ID', 'Notes'];
}

function appendAuditRow(sheet, data) {
  sheet.appendRow([
    data.auditTime,
    data.actionType,
    data.entryId,
    data.ccmOrChair,
    data.changedFields,
    data.reason,
    data.fromId,
    data.toId,
    data.notes
  ]);
}

function buildChangedFields(original, updated) {
  if (!original) return 'Full entry updated';
  var changes = [];
  var fields = ['reason', 'fromId', 'toId', 'notes'];
  fields.forEach(function(f) {
    if (original[f] !== updated[f]) {
      changes.push(f + ': "' + (original[f] || '') + '" → "' + (updated[f] || '') + '"');
    }
  });
  return changes.join(' | ');
}

function getCTTimestamp() {
  return Utilities.formatDate(new Date(), 'America/Chicago', 'MM/dd/yyyy hh:mm:ss a');
}

// ════════════════════════════════════════════════════════════════════════
//  ONE-TIME SETUP — run this once manually from the Apps Script editor
//  to initialize all sheet tabs
// ════════════════════════════════════════════════════════════════════════
function setupSheets() {
  var ss = SpreadsheetApp.openById(SS_ID);
  getOrCreateSheet(ss, SHEET_ROSTER,   ['Badge ID', 'Last Name', 'First Name', 'County', 'E-mail', 'Phone Number', 'Committee', 'Gender']);
  getOrCreateSheet(ss, SHEET_MASTER,   getMCLHeaders());
  getOrCreateSheet(ss, SHEET_QUORUM,   null);
  getOrCreateSheet(ss, SHEET_REASSIGN, null);
  getOrCreateSheet(ss, SHEET_AUDIT,    getAuditHeaders());
  refreshQuorumSheet(ss);
  refreshReassignmentSheet(ss);
  Logger.log('Setup complete. All sheets initialized.');
}
