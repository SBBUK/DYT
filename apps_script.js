// ── DYT Apps Script ──
// Paste into Extensions → Apps Script → Code.gs
// Deploy as Web App: Execute as Me, Anyone can access

const SHEET_ID = '1-Yb9u6h0v6BQffZqOhV7kjYXPtSjqqPWC0cyzgKxWCY';

const SESSIONS_HEADERS = ['ID', 'Name', 'Date', 'Time', 'Location', 'Max Spots', 'Status', 'Drop Date', 'Drop Time', 'Price'];
const REG_HEADERS = ['Timestamp', 'Name', 'Handle', 'First Session', 'Status', 'Email', 'Session_ID', 'Payment Intent'];

function getMilestone(sessions) {
  if (sessions >= 100) return 'Legendary';
  if (sessions >= 50)  return 'Full DYT Kit';
  if (sessions >= 25)  return 'DYT Shirt';
  if (sessions >= 10)  return 'DYT Family';
  return '';
}

function getNextMilestone(sessions) {
  if (sessions < 10)  return { name: 'DYT Family',   target: 10 };
  if (sessions < 25)  return { name: 'DYT Shirt',    target: 25 };
  if (sessions < 50)  return { name: 'Full DYT Kit', target: 50 };
  if (sessions < 100) return { name: 'Legendary',    target: 100 };
  return { name: 'Maxed Out', target: 100 };
}

function getMembersSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return sheet;
  }
  // Migrate: add any missing headers (e.g. new columns added after the sheet was first created)
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  headers.forEach((header, i) => {
    if (currentHeaders[i] !== header) {
      sheet.getRange(1, i + 1).setValue(header).setFontWeight('bold');
    }
  });
  return sheet;
}

function parseDropDateTime(dropDate, dropTime) {
  try {
    const parts = String(dropDate).split('/').map(Number);
    const day = parts[0], month = parts[1];
    const year = new Date().getFullYear();
    let h, min;
    const h24 = String(dropTime).match(/^(\d{1,2}):(\d{2})$/);
    if (h24) { h = parseInt(h24[1]); min = parseInt(h24[2]); }
    else {
      const h12 = String(dropTime).match(/(\d+):(\d+)\s*(AM|PM)/i);
      if (!h12) return null;
      h = parseInt(h12[1]); min = parseInt(h12[2]);
      const ampm = h12[3].toUpperCase();
      if (ampm === 'PM' && h !== 12) h += 12;
      if (ampm === 'AM' && h === 12) h = 0;
    }
    return new Date(year, month - 1, day, h, min, 0);
  } catch(e) { return null; }
}

function formatTime12h(timeStr) {
  const m = String(timeStr).match(/^(\d{1,2}):(\d{2})/);
  if (!m) return timeStr;
  let h = parseInt(m[1]);
  const min = m[2];
  const ampm = h >= 12 ? 'pm' : 'am';
  h = h % 12 || 12;
  return min === '00' ? `${h}${ampm}` : `${h}:${min}${ampm}`;
}

function isDropLive(sessionRow) {
  const status   = String(sessionRow[6] || 'soon').toLowerCase().trim();
  if (status === 'live') return true;
  if (status !== 'soon') return false;
  const dropDate = sessionRow[7] || '';
  const dropTime = sessionRow[8] || '';
  if (!dropDate || !dropTime) return false;
  const dt = parseDropDateTime(dropDate, dropTime);
  return dt && new Date() >= dt;
}

function isNoShow(handle) {
  const sheet = getOrCreateSheet('No-Show List', ['Handle', 'Session Missed', 'Date Added']);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toLowerCase().trim() === handle.toLowerCase().trim()) return true;
  }
  return false;
}

function getSessionRow(sessionId) {
  const sessionsSheet = getOrCreateSheet('Sessions', SESSIONS_HEADERS);
  const sessionsData  = sessionsSheet.getDataRange().getValues();
  for (let i = 1; i < sessionsData.length; i++) {
    if (String(sessionsData[i][0]).trim() === String(sessionId).trim()) return sessionsData[i];
  }
  return null;
}

// Marks a registration row with newStatus, and if it freed up a Confirmed
// spot, promotes the longest-waiting person on the waitlist for that
// session. Returns the promoted row (or null if no one was promoted).
function releaseSpot(regSheet, rowIndex, sessionId, newStatus, promote) {
  regSheet.getRange(rowIndex, 5).setValue(newStatus);
  if (!promote) return null;

  const data = regSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[6] || '').trim() === String(sessionId).trim() && row[4] === 'Waitlist') {
      regSheet.getRange(i + 1, 5).setValue('Confirmed');
      return row;
    }
  }
  return null;
}

function notifyPromotion(promotedRow, sessionId) {
  const sessionRow = getSessionRow(sessionId);
  const email = promotedRow[5];
  if (!sessionRow || !email) return;
  try {
    sendConfirmationEmail(email, 'Confirmed', sessionRow[1], sessionRow[2], sessionRow[3], sessionRow[4], false, promotedRow[3] === 'Yes');
  } catch (err) {
    Logger.log(err);
  }
}

// Runs fn(regData) under a script-wide lock and appends the resulting row
// (if any) before releasing it, so two simultaneous registrations can't
// both read the same "confirmed" count and overbook past maxSpots.
// fn must return either { duplicate: <row> } or { regStatus, confirmed, newRow }.
function withRegistrationLock(regSheet, fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return { error: true };
  }
  try {
    const result = fn(regSheet.getDataRange().getValues());
    if (!result.duplicate) regSheet.appendRow(result.newRow);
    return result;
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const action = e.parameter.action || '';

  if (action === 'sessions_list') {
    const sessionsSheet = getOrCreateSheet('Sessions', SESSIONS_HEADERS);
    const regSheet      = getOrCreateSheet('Session Registrations', REG_HEADERS);
    const sessionsData  = sessionsSheet.getDataRange().getValues();
    const regData       = regSheet.getDataRange().getValues();
    const confirmedMap  = {};
    for (let i = 1; i < regData.length; i++) {
      const row = regData[i];
      if (!row[0]) continue;
      const sid = String(row[6] || '').trim();
      const status = String(row[4] || '').trim();
      if (sid && status === 'Confirmed') confirmedMap[sid] = (confirmedMap[sid] || 0) + 1;
    }
    const sessions = [];
    for (let i = 1; i < sessionsData.length; i++) {
      const row = sessionsData[i];
      if (!row[0]) continue;
      const id = String(row[0]).trim();
      const maxSpots = parseInt(row[5] || 10);
      const confirmed = confirmedMap[id] || 0;
      sessions.push({ id, name: row[1] || '', date: row[2] || '', time: row[3] || '', location: row[4] || '', maxSpots, confirmed, spotsLeft: Math.max(0, maxSpots - confirmed), status: String(row[6] || 'soon').toLowerCase().trim(), dropDate: row[7] || '', dropTime: row[8] || '', price: parseFloat(row[9] || 0) });
    }
    return response(sessions);
  }

  if (action === 'session_info') {
    const sessionId = e.parameter.id || null;
    if (sessionId) {
      const sessionsSheet = getOrCreateSheet('Sessions', SESSIONS_HEADERS);
      const regSheet      = getOrCreateSheet('Session Registrations', REG_HEADERS);
      const sessionsData  = sessionsSheet.getDataRange().getValues();
      let sessionRow = null;
      for (let i = 1; i < sessionsData.length; i++) {
        if (String(sessionsData[i][0]).trim() === String(sessionId).trim()) { sessionRow = sessionsData[i]; break; }
      }
      if (!sessionRow) return response({ error: 'Session not found', isActive: false });
      const maxSpots  = parseInt(sessionRow[5] || 10);
      const status    = String(sessionRow[6] || 'soon').toLowerCase().trim();
      const isActive  = isDropLive(sessionRow);
      const regData   = regSheet.getDataRange().getValues();
      const confirmed = regData.filter((row, i) => i > 0 && String(row[6]).trim() === String(sessionId).trim() && row[4] === 'Confirmed').length;
      return response({ isActive, id: String(sessionRow[0]).trim(), sessionName: sessionRow[1] || 'DYT Session', date: sessionRow[2] || '', time: sessionRow[3] || '', location: sessionRow[4] || 'Roehampton Sport & Fitness Centre', maxSpots, confirmed, spotsLeft: Math.max(0, maxSpots - confirmed), status, price: parseFloat(sessionRow[9] || 0) });
    }
    const configSheet = getOrCreateSheet('Session Config', ['Field', 'Value']);
    const regSheet    = getOrCreateSheet('Session Registrations', REG_HEADERS);
    const config = {};
    configSheet.getDataRange().getValues().forEach((row, i) => { if (i > 0 && row[0]) config[row[0]] = row[1]; });
    const maxSpots  = parseInt(config['max_spots'] || 10);
    const isActive  = config['is_active'] === true || config['is_active'] === 'TRUE' || config['is_active'] === 'true';
    const regData   = regSheet.getDataRange().getValues();
    const confirmed = regData.filter((row, i) => i > 0 && row[4] === 'Confirmed').length;
    return response({ isActive, sessionName: config['session_name'] || 'DYT Session', date: config['date'] || '', time: config['time'] || '', location: config['location'] || 'Roehampton Sport & Fitness Centre', maxSpots, confirmed, spotsLeft: Math.max(0, maxSpots - confirmed) });
  }

  if (action === 'session_register') {
    const name          = e.parameter.name    || '';
    const handle        = ('@' + (e.parameter.handle || '').replace(/^@/, '')).toLowerCase();
    const first         = e.parameter.first === 'true';
    const email         = e.parameter.email  || '';
    const sessionId     = e.parameter.id     || null;
    const paymentIntent = e.parameter.payment_intent || '';
    if (!name || !handle || handle === '@') return response({ status: 'error', message: 'Missing name or handle' });

    const noShow = isNoShow(handle);
    const isActiveReg = (status) => status === 'Confirmed' || status === 'Waitlist';

    const regSheet = getOrCreateSheet('Session Registrations', REG_HEADERS);
    if (sessionId) {
      const sessionsSheet = getOrCreateSheet('Sessions', SESSIONS_HEADERS);
      const sessionsData  = sessionsSheet.getDataRange().getValues();
      let sessionRow = null;
      for (let i = 1; i < sessionsData.length; i++) {
        if (String(sessionsData[i][0]).trim() === String(sessionId).trim()) { sessionRow = sessionsData[i]; break; }
      }
      if (!sessionRow) return response({ status: 'error', message: 'Session not found' });
      if (!isDropLive(sessionRow)) {
        const dropTime = formatTime12h(sessionRow[8] || '');
        return response({ status: 'not_open', message: dropTime ? `Come back at ${dropTime}!` : 'Sign-ups open soon!' });
      }
      const maxSpots = parseInt(sessionRow[5] || 10);
      const result = withRegistrationLock(regSheet, (regData) => {
        const duplicate = regData.slice(1).find(row => (row[2] || '').toLowerCase() === handle && String(row[6] || '').trim() === String(sessionId).trim() && isActiveReg(row[4]));
        if (duplicate) return { duplicate };
        const confirmed = regData.filter((row, i) => i > 0 && String(row[6]).trim() === String(sessionId).trim() && row[4] === 'Confirmed').length;
        const regStatus = (noShow || confirmed >= maxSpots) ? 'Waitlist' : 'Confirmed';
        return { regStatus, confirmed, newRow: [new Date(), name, handle, first ? 'Yes' : 'No', regStatus, email, sessionId, paymentIntent] };
      });
      if (result.error) return response({ status: 'error', message: 'Busy right now — please try again in a moment.' });
      if (result.duplicate) return response({ status: result.duplicate[4] === 'Confirmed' ? 'confirmed' : 'waitlist', duplicate: true });
      try { if (email) sendConfirmationEmail(email, result.regStatus, sessionRow[1], sessionRow[2], sessionRow[3], sessionRow[4], noShow, first); } catch(err) { Logger.log(err); }
      return response({ status: result.regStatus.toLowerCase(), spotsLeft: Math.max(0, maxSpots - result.confirmed - 1), noShow });
    }
    const configSheet = getOrCreateSheet('Session Config', ['Field', 'Value']);
    const config = {};
    configSheet.getDataRange().getValues().forEach((row, i) => { if (i > 0 && row[0]) config[row[0]] = row[1]; });
    const isActive = config['is_active'] === true || config['is_active'] === 'TRUE' || config['is_active'] === 'true';
    if (!isActive) return response({ status: 'error', message: 'No active session' });
    const maxSpots = parseInt(config['max_spots'] || 10);
    const result = withRegistrationLock(regSheet, (regData) => {
      const duplicate = regData.slice(1).find(row => (row[2] || '').toLowerCase() === handle && isActiveReg(row[4]));
      if (duplicate) return { duplicate };
      const confirmed = regData.filter((row, i) => i > 0 && row[4] === 'Confirmed').length;
      const regStatus = (noShow || confirmed >= maxSpots) ? 'Waitlist' : 'Confirmed';
      return { regStatus, confirmed, newRow: [new Date(), name, handle, first ? 'Yes' : 'No', regStatus, email, '', paymentIntent] };
    });
    if (result.error) return response({ status: 'error', message: 'Busy right now — please try again in a moment.' });
    if (result.duplicate) return response({ status: result.duplicate[4] === 'Confirmed' ? 'confirmed' : 'waitlist', duplicate: true });
    try { if (email) sendConfirmationEmail(email, result.regStatus, config['session_name'], config['date'], config['time'], config['location'], noShow, first); } catch(err) { Logger.log(err); }
    return response({ status: result.regStatus.toLowerCase(), spotsLeft: Math.max(0, maxSpots - result.confirmed - 1), noShow });
  }

  if (action === 'cancel_spot') {
    const handle    = ('@' + (e.parameter.handle || '').replace(/^@/, '')).toLowerCase();
    const sessionId = e.parameter.session || null;
    if (!handle || handle === '@') return response({ status: 'error', message: 'Missing handle' });

    const regSheet = getOrCreateSheet('Session Registrations', REG_HEADERS);
    const data      = regSheet.getDataRange().getValues();

    let activeRow = -1, alreadyCancelled = false;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if ((row[2] || '').toLowerCase().trim() !== handle) continue;
      if (sessionId && String(row[6] || '').trim() !== String(sessionId).trim()) continue;
      if (row[4] === 'Confirmed' || row[4] === 'Waitlist') { activeRow = i; break; }
      if (row[4] === 'Cancelled' || row[4] === 'Refunded') alreadyCancelled = true;
    }

    if (activeRow === -1) {
      return response({ status: alreadyCancelled ? 'already_cancelled' : 'not_found' });
    }

    const row          = data[activeRow];
    const wasConfirmed = row[4] === 'Confirmed';
    const rowSessionId = String(row[6] || '').trim();

    const promoted = releaseSpot(regSheet, activeRow + 1, rowSessionId, 'Cancelled', wasConfirmed);
    if (promoted) notifyPromotion(promoted, rowSessionId);

    return response({ status: 'cancelled' });
  }

  if (action === 'session_refund') {
    const paymentIntent = (e.parameter.payment_intent || '').trim();
    if (!paymentIntent) return response({ status: 'error', message: 'Missing payment_intent' });

    const regSheet = getOrCreateSheet('Session Registrations', REG_HEADERS);
    const data      = regSheet.getDataRange().getValues();

    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][7] || '').trim() === paymentIntent) { targetRow = i; break; }
    }
    if (targetRow === -1) return response({ status: 'not_found' });

    const row = data[targetRow];
    if (row[4] === 'Cancelled' || row[4] === 'Refunded') return response({ status: 'already_cancelled' });

    const wasConfirmed = row[4] === 'Confirmed';
    const sessionId    = String(row[6] || '').trim();
    const email        = row[5];

    const promoted = releaseSpot(regSheet, targetRow + 1, sessionId, 'Refunded', wasConfirmed);
    if (promoted) notifyPromotion(promoted, sessionId);

    const sessionRow = getSessionRow(sessionId);
    if (sessionRow) {
      try { sendRefundEmail(email, sessionRow[1], sessionRow[2], sessionRow[3], sessionRow[9]); } catch (err) { Logger.log(err); }
    }

    return response({ status: 'refunded' });
  }

  if (action === 'all') {
    const sheet = getMembersSheet();
    const data  = sheet.getDataRange().getValues();
    const members = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      members.push({ name: data[i][0], handle: data[i][1], sessions: data[i][2] || 0, milestone: data[i][4] || '' });
    }
    return response(members);
  }

  if (action === 'lookup') {
    const handle = (e.parameter.handle || '').toLowerCase().replace('@', '');
    const sheet  = getMembersSheet();
    const data   = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      if ((row[1] || '').toLowerCase().replace('@', '') === handle) {
        const sessions = row[2] || 0;
        return response({ found: true, name: row[0], handle: row[1], sessions, milestone: row[4] || '', next: getNextMilestone(sessions) });
      }
    }
    return response({ found: false });
  }

  return response({ error: 'Unknown action' });
}

function doPost(e) {
  const body  = JSON.parse(e.postData.contents);
  const sheet = getMembersSheet();
  const data  = sheet.getDataRange().getValues();
  if (body.action === 'checkin') {
    const handles  = (body.handles || []).map(h => h.toLowerCase().replace('@', ''));
    const achieved = [];
    for (let i = 1; i < data.length; i++) {
      const rowHandle = (data[i][1] || '').toLowerCase().replace('@', '');
      if (!rowHandle || !handles.includes(rowHandle)) continue;
      const oldCount     = data[i][2] || 0;
      const newCount     = oldCount + 1;
      const oldMilestone = data[i][4] || '';
      const newMilestone = getMilestone(newCount);
      sheet.getRange(i + 1, 3).setValue(newCount);
      if (newMilestone && newMilestone !== oldMilestone) {
        sheet.getRange(i + 1, 5).setValue(newMilestone);
        achieved.push({ name: data[i][0], handle: data[i][1], sessions: newCount, milestone: newMilestone });
      }
    }
    return response({ success: true, milestones: achieved });
  }
  if (body.action === 'register') {
    sheet.appendRow([body.name, body.handle, 0, new Date(), '']);
    return response({ success: true });
  }
  return response({ error: 'Unknown action' });
}

function sendConfirmationEmail(email, status, sName, sDate, sTime, sLoc, noShow, first) {
  if (status === 'Confirmed') {
    MailApp.sendEmail({
      to: email,
      name: sName || "DYT Breakfast Club",
      subject: "You're In — " + sName + " · " + sDate,
      body: (first ? "You're in. Welcome to the family." : "You're in. Welcome back, family.") + "\n\nSession: " + sName + "\nDate: " + sDate + "\nTime: " + sTime + "\nLocation: " + sLoc + "\n\nCan't make it? Give 12 hours notice.\nNo-shows sit out the next drop. No exceptions.\n\nFor Hoopers. By Hoopers. DYT Family."
    });
  } else if (noShow) {
    MailApp.sendEmail({
      to: email,
      name: sName || "DYT Breakfast Club",
      subject: "Waitlist — " + sName + " · " + sDate,
      body: "You're on the waitlist for this drop.\n\nYou previously missed a session without giving 12 hours notice.\nAs a result you've been moved to the waitlist for this drop.\n\nIf a spot opens up we'll be in touch.\nAttend this session and you'll be fully reinstated for future drops.\n\nFor Hoopers. By Hoopers. DYT Family."
    });
  } else {
    MailApp.sendEmail({
      to: email,
      name: sName || "DYT Breakfast Club",
      subject: "Waitlist — " + sName + " · " + sDate,
      body: "All spots are taken.\n\nYou're on the waitlist for " + sName + " on " + sDate + " at " + sTime + ".\n\nIf a spot opens up we'll be in touch as soon as possible.\n\nFor Hoopers. By Hoopers. DYT Family."
    });
  }
}

function sendRefundEmail(email, sName, sDate, sTime, price) {
  if (!email) return;
  MailApp.sendEmail({
    to: email,
    name: sName || "DYT Breakfast Club",
    subject: "Refund Processed — " + sName + " · " + sDate,
    body: "Your payment for " + sName + " on " + sDate + " at " + sTime + " has been refunded" + (price ? " (£" + price + ")" : "") + ".\n\nYour spot has been released. If you'd still like to come, you're welcome to book another session.\n\nFor Hoopers. By Hoopers. DYT Family."
  });
}

function testEmail() {
  MailApp.sendEmail({ to: 'team@simplybritishballers.com', subject: 'DYT Test Email', body: 'Email is working.', name: 'DYT Breakfast Club' });
}

function response(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
