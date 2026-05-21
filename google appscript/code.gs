// ════════════════════════════════════════════════════════
//  Manav Rahat — Member Entry System
//  Google Apps Script Backend — FINAL
// ════════════════════════════════════════════════════════

const MEMBERS_SHEET = 'master_data';
const ENTRIES_SHEET = 'transaction_data';
const USERS_SHEET     = 'user_master';
const DUP_SLIP_SHEET  = 'duplicateslip';

// duplicateslip columns (0-based):
// A=0 Entry id | B=1 Member_id | C=2 Member Name | D=3 Manav_Rahat
// E=4 remark | F=5 Userid | G=6 Created on ( Date & time )

// user_master columns (0-based):
// A=0 userid | B=1 password | C=2 name | D=3 range_start
// E=4 range_end | F=5 active | G=6 is_admin  ← NEW

// ════════════════════════════════════════════════════════
//  ENTRY POINT
// ════════════════════════════════════════════════════════
function doGet(e) {
  const p = e.parameter, action = p.action || '';
  try {
    switch (action) {
      case 'login':          return jsonResponse(loginUser(p.userid, p.password));
      case 'getMembers':     return jsonResponse(getMembers());
      case 'checkDup':       return jsonResponse(checkDuplicate(p.member_id));
      case 'addEntry':       return jsonResponse(addEntry(p));
      case 'getEntries':     return jsonResponse(getEntries(p.logged_by));        // my entries
      case 'getAllEntries':   return jsonResponse(getAllEntries(p.logged_by));     // admin: by user filter
      case 'getSummary':     return jsonResponse(getSummary());                   // admin: summary
      case 'getUserList':    return jsonResponse(getUserList());                  // admin: user dropdown
      case 'updateEntry':    return jsonResponse(updateEntry(p));               // admin: edit entry
      case 'checkDupSlip':   return jsonResponse(checkDupSlipMember(p.member_id)); // dup slip: validate member
      case 'addDupSlip':     return jsonResponse(addDupSlip(p));                   // dup slip: save record
      case 'getDupSlipById': return jsonResponse(getDupSlipById(p.member_id));     // dup slip: get by member
      default:               return jsonResponse({ error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ════════════════════════════════════════════════════════
//  loginUser
//  Returns: user info + all members + is_admin flag
//  next_entry_id is computed from existing entries so
//  it always picks up from where user left off (survives
//  multiple login/logout cycles)
// ════════════════════════════════════════════════════════
function loginUser(userid, password) {
  if (!userid || !password)
    return { success: false, error: 'userid and password are required' };

  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(USERS_SHEET);
  if (!userSheet) return { success: false, error: 'Sheet "' + USERS_SHEET + '" not found.' };

  const rows = userSheet.getDataRange().getValues();
  let userRow = null;

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (String(r[0]).trim().toLowerCase() === userid.trim().toLowerCase()) {
      if (String(r[5]).trim().toUpperCase() === 'NO')
        return { success: false, error: 'Account is disabled. Contact administrator.' };
      if (String(r[1]).trim() !== password.trim())
        return { success: false, error: 'Incorrect password.' };
      userRow = r; break;
    }
  }
  if (!userRow) return { success: false, error: 'User not found.' };

  const rangeStart = Number(userRow[3]);
  const rangeEnd   = Number(userRow[4]);
  const isAdmin    = String(userRow[6] || '').trim().toUpperCase() === 'YES';

  // next_entry_id — always computed fresh from sheet
  const nextId = getNextEntryId(ss, rangeStart, rangeEnd);

  // Members NOT included here — fetched separately via getMembers action
  // This keeps login response small and fast
  return {
    success:       true,
    userid:        String(userRow[0]).trim(),
    name:          String(userRow[2]).trim(),
    range_start:   rangeStart,
    range_end:     rangeEnd,
    next_entry_id: nextId,
    is_admin:      isAdmin,
  };
}

function loadAllMembers(ss) {
  const sheet = ss.getSheetByName(MEMBERS_SHEET);
  if (!sheet) return [];
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const data = sheet.getRange(2, 1, last - 1, 7).getValues();
  const out  = [];
  for (let i = 0; i < data.length; i++) {
    const id = String(data[i][0]).trim().toUpperCase();  // normalize to UPPERCASE
    if (!id) continue;
    const mr = String(data[i][6]).trim().toUpperCase();
    out.push({
      memberid:              id,  // always uppercase — matches mmap key
      new_code:              String(data[i][1]).trim(),
      name:                  String(data[i][2]).trim(),
      address:               String(data[i][3]).trim(),
      pincode:               String(data[i][4]).trim(),
      mobile:                String(data[i][5]).trim(),
      manav_rahat:           String(data[i][6]).trim(),
      is_marksheet_eligible: mr === 'M',
    });
  }
  return out;
}

// ════════════════════════════════════════════════════════
//  getMembers — separate action, returns all members
//  Called right after login in the browser
// ════════════════════════════════════════════════════════
function getMembers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const members = loadAllMembers(ss);
  return { success: true, members: members, count: members.length };
}

// ════════════════════════════════════════════════════════
//  checkDuplicate — called at scan time, checks TODAY
//  across ALL users
// ════════════════════════════════════════════════════════
function checkDuplicate(memberId) {
  if (!memberId) return { duplicate: false };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ENTRIES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { duplicate: false };

  const last  = sheet.getLastRow() - 1;
  const mIds  = sheet.getRange(2, 2, last, 1).getValues(); // col B member_id
  const uIds  = sheet.getRange(2, 7, last, 1).getValues(); // col G userid
  const times = sheet.getRange(2, 8, last, 1).getValues(); // col H timestamp

  const query = String(memberId).trim().toUpperCase();
  const today = new Date();
  const tY = today.getFullYear(), tM = today.getMonth(), tD = today.getDate();

  for (let i = 0; i < mIds.length; i++) {
    if (String(mIds[i][0]).trim().toUpperCase() !== query) continue;
    const ts = times[i][0];
    if (!ts) continue;
    const d = new Date(ts);
    if (d.getFullYear() === tY && d.getMonth() === tM && d.getDate() === tD) {
      return {
        duplicate: true,
        by: String(uIds[i][0]).trim(),
        at: d.toLocaleTimeString('en-IN', { hour:'2-digit', minute:'2-digit', hour12:true }),
      };
    }
  }
  return { duplicate: false };
}

// ════════════════════════════════════════════════════════
//  addEntry — ZERO reads, only appendRow → fast
// ════════════════════════════════════════════════════════
function addEntry(p) {
  const req = ['memberid','member_name','qty','entry_method','logged_by','entry_id'];
  for (const f of req)
    if (!p[f] || String(p[f]).trim() === '')
      return { success: false, error: 'Missing: ' + f };

  const marksheet = String(p.marksheet).trim() === 'Yes' ? 'Yes' : 'No';
  let amount = 0;
  if (marksheet !== 'Yes') {
    const qty = Number(p.qty);
    if      (qty === 30) amount = p.entry_method === 'qrcode' ? 500 : 600;
    else if (qty === 20) amount = 400;
    else if (qty === 10) amount = 200;
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ENTRIES_SHEET);
  if (!sheet) return { success: false, error: 'Sheet "' + ENTRIES_SHEET + '" not found.' };

  if (sheet.getLastRow() === 0)
    sheet.appendRow(['Entry id','Member_id','Member Name','Manav_Rahat','Qty','Amount','Userid','Created on ( Date & time )','Marksheet','changedby','updatedon']);

  const now = new Date();
  sheet.appendRow([
    Number(p.entry_id), p.memberid, p.member_name,
    p.manav_rahat||'', Number(p.qty), amount,
    p.logged_by, now, marksheet,
  ]);

  return { success:true, entry_id:Number(p.entry_id), amount, marksheet, timestamp:now.toISOString() };
}

// ════════════════════════════════════════════════════════
//  getEntries — my entries (filtered by logged_by)
// ════════════════════════════════════════════════════════
function getEntries(loggedBy) {
  if (!loggedBy) return { entries:[] };
  const rows = readEntrySheet();
  const target = loggedBy.toLowerCase();
  return {
    entries: rows.filter(r => r.userid.toLowerCase() === target)
  };
}

// ════════════════════════════════════════════════════════
//  getAllEntries — admin: all entries, optional user filter
//  p.logged_by = specific userid OR 'ALL' for everyone
// ════════════════════════════════════════════════════════
function getAllEntries(loggedBy) {
  const rows = readEntrySheet();
  if (!loggedBy || loggedBy.toUpperCase() === 'ALL') {
    return { entries: rows };
  }
  return {
    entries: rows.filter(r => r.userid.toLowerCase() === loggedBy.toLowerCase())
  };
}

// ════════════════════════════════════════════════════════
//  getSummary — admin: user-wise aggregation
//  Returns array of { userid, name, count, total_qty, total_amount, marksheet_count }
// ════════════════════════════════════════════════════════
function getSummary() {
  const rows = readEntrySheet();

  // Build user display name map from user_master
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(USERS_SHEET);
  const nameMap   = {};
  if (userSheet) {
    const ud = userSheet.getDataRange().getValues();
    for (let i = 1; i < ud.length; i++) {
      nameMap[String(ud[i][0]).trim().toLowerCase()] = String(ud[i][2]).trim();
    }
  }

  // Aggregate
  const agg = {};
  for (const r of rows) {
    const uid = r.userid.toLowerCase();
    if (!agg[uid]) {
      agg[uid] = {
        userid:          r.userid,
        name:            nameMap[uid] || r.userid,
        count:           0,
        total_qty:       0,
        total_amount:    0,
        marksheet_count: 0,
      };
    }
    agg[uid].count++;
    agg[uid].total_qty    += Number(r.qty || 0);
    agg[uid].total_amount += Number(r.amount || 0);
    if (r.marksheet === 'Yes') agg[uid].marksheet_count++;
  }

  // Overall totals row
  const summary = Object.values(agg);
  const grand = {
    userid: '__TOTAL__',
    name:   'GRAND TOTAL',
    count:           summary.reduce((s,r)=>s+r.count,0),
    total_qty:       summary.reduce((s,r)=>s+r.total_qty,0),
    total_amount:    summary.reduce((s,r)=>s+r.total_amount,0),
    marksheet_count: summary.reduce((s,r)=>s+r.marksheet_count,0),
  };

  return { summary, grand };
}

// ════════════════════════════════════════════════════════
//  getUserList — admin: returns list of non-admin users
//  for the dropdown in Report 2
// ════════════════════════════════════════════════════════
function getUserList() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName(USERS_SHEET);
  if (!userSheet) return { users: [] };

  const rows  = userSheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < rows.length; i++) {
    const r      = rows[i];
    const active = String(r[5]).trim().toUpperCase();
    if (active === 'NO') continue;
    users.push({
      userid: String(r[0]).trim(),
      name:   String(r[2]).trim(),
    });
  }
  return { users };
}

// ════════════════════════════════════════════════════════
//  updateEntry — admin only
//  Finds row by entry_id (col A) and updates editable fields.
//  Member ID (col B) is NEVER changed.
//  Editable: member_name(C), qty(E), entry_method(implicit via amount),
//            amount(F), userid(G), marksheet(I)
//  Amount is re-calculated server-side from qty + entry_method.
// ════════════════════════════════════════════════════════
function updateEntry(p) {
  if (!p.entry_id) return { success: false, error: 'entry_id required' };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ENTRIES_SHEET);
  if (!sheet) return { success: false, error: 'Entries sheet not found' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, error: 'No entries found' };

  // Find the row with matching entry_id (col A)
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let targetRow = -1;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(p.entry_id).trim()) {
      targetRow = i + 2; // +2 because data starts at row 2 (1-based)
      break;
    }
  }
  if (targetRow === -1) return { success: false, error: 'Entry ID ' + p.entry_id + ' not found' };

  // Server-side amount re-calculation
  const marksheet = String(p.marksheet||'No').trim() === 'Yes' ? 'Yes' : 'No';
  let amount = 0;
  if (marksheet !== 'Yes') {
    const qty = Number(p.qty);
    if      (qty === 30) amount = p.entry_method === 'qrcode' ? 500 : 600;
    else if (qty === 20) amount = 400;
    else if (qty === 10) amount = 200;
    else amount = Number(p.amount) || 0; // fallback to passed amount
  }

  // Update only editable columns — col B (member_id) is NEVER touched
  // C=3: Member Name, E=5: Qty, F=6: Amount, G=7: Userid, I=9: Marksheet
  // J=10: changedby (admin who edited), K=11: updatedon (timestamp of edit)
  sheet.getRange(targetRow, 3).setValue(p.member_name  || '');  // C: Member Name
  sheet.getRange(targetRow, 5).setValue(Number(p.qty)  || 0);   // E: Qty
  sheet.getRange(targetRow, 6).setValue(amount);                 // F: Amount
  sheet.getRange(targetRow, 7).setValue(p.userid       || '');  // G: Userid
  sheet.getRange(targetRow, 9).setValue(marksheet);              // I: Marksheet
  sheet.getRange(targetRow, 10).setValue(p.changedby   || '');  // J: changedby
  sheet.getRange(targetRow, 11).setValue(new Date());            // K: updatedon

  return { success: true, entry_id: p.entry_id, amount, marksheet, changedby: p.changedby, updatedon: new Date().toISOString() };
}

// ════════════════════════════════════════════════════════
//  SHARED HELPERS
// ════════════════════════════════════════════════════════

// Read all entries from transaction_data — shared by multiple functions
function readEntrySheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ENTRIES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data    = sheet.getRange(2, 1, sheet.getLastRow()-1, 11).getValues();
  const entries = [];
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (!r[0]) continue; // skip blank rows
    entries.push({
      entry_id:    r[0],
      member_id:   String(r[1]).trim(),
      member_name: String(r[2]).trim(),
      manav_rahat: String(r[3]).trim(),
      qty:         r[4],
      amount:      r[5],
      userid:      String(r[6]).trim(),
      created_on:  r[7] ? new Date(r[7]).toISOString() : '',
      marksheet:   String(r[8]||'No').trim(),
      changedby:   String(r[9]||'').trim(),
      changed_on:  r[10] ? new Date(r[10]).toISOString() : '',
    });
  }
  return entries;
}

// Compute next entry ID for a user — always reads fresh from sheet
// so it survives multiple login/logout cycles correctly
function getNextEntryId(ss, rangeStart, rangeEnd) {
  const sheet = ss.getSheetByName(ENTRIES_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return rangeStart;
  const data  = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
  let maxId   = rangeStart - 1;
  for (const row of data) {
    const id = Number(row[0]);
    if (id >= rangeStart && id <= rangeEnd && id > maxId) maxId = id;
  }
  const next = maxId + 1;
  if (next > rangeEnd) throw new Error('Entry ID range exhausted for this user');
  return next;
}

// ════════════════════════════════════════════════════════
//  checkDupSlipMember
//  1. Check duplicateslip sheet — if already exists, block and show record
//  2. If not in dup slip, check transaction_data — show warning if found
// ════════════════════════════════════════════════════════
function checkDupSlipMember(memberId) {
  if (!memberId) return { status: 'error', error: 'member_id required' };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const query = String(memberId).trim().toUpperCase();

  // --- Check 1: duplicateslip sheet ---
  const dupSheet = ss.getSheetByName(DUP_SLIP_SHEET);
  if (dupSheet && dupSheet.getLastRow() > 1) {
    const dupData = dupSheet.getRange(2, 1, dupSheet.getLastRow()-1, 7).getValues();
    for (let i = 0; i < dupData.length; i++) {
      const row = dupData[i];
      if (String(row[1]).trim().toUpperCase() === query) {
        // Already has a dup slip — BLOCK
        return {
          status:      'dup_slip_exists',
          entry_id:    row[0],
          member_id:   row[1],
          member_name: row[2],
          remark:      row[4],
          userid:      row[5],
          created_on:  row[6] ? new Date(row[6]).toISOString() : '',
          message:     'Duplicate slip already issued for this member.',
        };
      }
    }
  }

  // --- Check 2: transaction_data sheet ---
  const txSheet = ss.getSheetByName(ENTRIES_SHEET);
  if (txSheet && txSheet.getLastRow() > 1) {
    const txData = txSheet.getRange(2, 1, txSheet.getLastRow()-1, 9).getValues();
    for (let i = 0; i < txData.length; i++) {
      const row = txData[i];
      if (String(row[1]).trim().toUpperCase() === query) {
        // Found in transaction_data — show warning but allow
        return {
          status:      'entry_exists',
          entry_id:    row[0],
          member_id:   row[1],
          member_name: row[2],
          qty:         row[4],
          amount:      row[5],
          userid:      row[6],
          created_on:  row[7] ? new Date(row[7]).toISOString() : '',
          marksheet:   row[8] || 'No',
          message:     'Entry already exists in transaction data.',
        };
      }
    }
  }

  // --- All clear ---
  return { status: 'ok' };
}

// ════════════════════════════════════════════════════════
//  addDupSlip
//  Saves a new row to duplicateslip sheet
//  Entry ID is auto-incremented (simple max+1, single user)
// ════════════════════════════════════════════════════════
function addDupSlip(p) {
  const required = ['member_id', 'member_name', 'userid'];
  for (const f of required)
    if (!p[f] || String(p[f]).trim() === '')
      return { success: false, error: 'Missing: ' + f };

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(DUP_SLIP_SHEET);

  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(DUP_SLIP_SHEET);
  }

  // Write header if sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Entry id', 'Member_id', 'Member Name', 'Manav_Rahat', 'remark', 'Userid', 'Created on ( Date & time )']);
    // Format header row
    sheet.getRange(1,1,1,7).setFontWeight('bold').setBackground('#1a3060').setFontColor('#ffffff');
  }

  // Get next entry ID (simple max+1 — single user, no range needed)
  let nextId = 1;
  if (sheet.getLastRow() > 1) {
    const ids = sheet.getRange(2, 1, sheet.getLastRow()-1, 1).getValues();
    let maxId = 0;
    for (const row of ids) {
      const id = Number(row[0]);
      if (id > maxId) maxId = id;
    }
    nextId = maxId + 1;
  }

  const now = new Date();
  sheet.appendRow([
    nextId,
    String(p.member_id).trim().toUpperCase(),
    p.member_name,
    p.manav_rahat || '',
    p.remark      || '',
    p.userid,
    now,
  ]);

  return { success: true, entry_id: nextId, timestamp: now.toISOString() };
}

// ════════════════════════════════════════════════════════
//  getDupSlipById — fetch a dup slip record by member_id
// ════════════════════════════════════════════════════════
function getDupSlipById(memberId) {
  if (!memberId) return { found: false };
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DUP_SLIP_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { found: false };

  const query = String(memberId).trim().toUpperCase();
  const data  = sheet.getRange(2, 1, sheet.getLastRow()-1, 7).getValues();
  for (const row of data) {
    if (String(row[1]).trim().toUpperCase() === query) {
      return {
        found:       true,
        entry_id:    row[0],
        member_id:   row[1],
        member_name: row[2],
        manav_rahat: row[3],
        remark:      row[4],
        userid:      row[5],
        created_on:  row[6] ? new Date(row[6]).toISOString() : '',
      };
    }
  }
  return { found: false };
}

function jsonResponse(data) {
  const out = ContentService.createTextOutput(JSON.stringify(data));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}
