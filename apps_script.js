// ── DYT Apps Script ──
// Paste into Extensions → Apps Script → Code.gs
// Deploy as Web App: Execute as Me, Anyone can access

const SHEET_ID = '1-Yb9u6h0v6BQffZqOhV7kjYXPtSjqqPWC0cyzgKxWCY';

function getMilestone(sessions) {
  if (sessions >= 100) return 'Legendary';
  if (sessions >= 50)  return 'Elite — Full Kit';
  if (sessions >= 25)  return 'SIMPLY Drop';
  if (sessions >= 10)  return 'DYT Family';
  return '';
}

function getNextMilestone(sessions) {
  if (sessions < 10)  return { name: 'DYT Family',       target: 10 };
  if (sessions < 25)  return { name: 'SIMPLY Drop',      target: 25 };
  if (sessions < 50)  return { name: 'Elite — Full Kit', target: 50 };
  if (sessions < 100) return { name: 'Legendary',        target: 100 };
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
  }
  return sheet;
}

function isNoShow(handle) {
  const sheet = getOrCreateSheet('No-Show List', ['Handle', 'Session Missed', 'Date Added']);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toLowerCase().trim() === handle.toLowerCase().trim()) return true;
  }
  return false;
}

function doGet(e) {
  const action = e.parameter.action || '';

  if (action === 'sessions_list') {
    const sessionsSheet = getOrCreateSheet('Sessions', ['ID', 'Name', 'Date', 'Time', 'Location', 'Max Spots', 'Status', 'Drop Date', 'Drop Time']);
    const regSheet      = getOrCreateSheet('Session Registrations', ['Timestamp', 'Name', 'Handle', 'First Session', 'Status', 'Email', 'Session_ID']);
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
      sessions.push({ id, name: row[1] || '', date: row[2] || '', time: row[3] || '', location: row[4] || '', maxSpots, confirmed, spotsLeft: Math.max(0, maxSpots - confirmed), status: String(row[6] || 'soon').toLowerCase().trim(), dropDate: row[7] || '', dropTime: row[8] || '' });
    }
    return response(sessions);
  }

  if (action === 'session_info') {
    const sessionId = e.parameter.id || null;
    if (sessionId) {
      const sessionsSheet = getOrCreateSheet('Sessions', ['ID', 'Name', 'Date', 'Time', 'Location', 'Max Spots', 'Status', 'Drop Date', 'Drop Time']);
      const regSheet      = getOrCreateSheet('Session Registrations', ['Timestamp', 'Name', 'Handle', 'First Session', 'Status', 'Email', 'Session_ID']);
      const sessionsData  = sessionsSheet.getDataRange().getValues();
      let sessionRow = null;
      for (let i = 1; i < sessionsData.length; i++) {
        if (String(sessionsData[i][0]).trim() === String(sessionId).trim()) { sessionRow = sessionsData[i]; break; }
      }
      if (!sessionRow) return response({ error: 'Session not found', isActive: false });
      const maxSpots  = parseInt(sessionRow[5] || 10);
      const status    = String(sessionRow[6] || 'soon').toLowerCase().trim();
      const isActive  = status === 'live';
      const regData   = regSheet.getDataRange().getValues();
      const confirmed = regData.filter((row, i) => i > 0 && String(row[6]).trim() === String(sessionId).trim() && row[4] === 'Confirmed').length;
      return response({ isActive, id: String(sessionRow[0]).trim(), sessionName: sessionRow[1] || 'DYT Session', date: sessionRow[2] || '', time: sessionRow[3] || '', location: sessionRow[4] || 'Roehampton Sport & Fitness Centre', maxSpots, confirmed, spotsLeft: Math.max(0, maxSpots - confirmed), status });
    }
    const configSheet = getOrCreateSheet('Session Config', ['Field', 'Value']);
    const regSheet    = getOrCreateSheet('Session Registrations', ['Timestamp', 'Name', 'Handle', 'First Session', 'Status', 'Email', 'Session_ID']);
    const config = {};
    configSheet.getDataRange().getValues().forEach((row, i) => { if (i > 0 && row[0]) config[row[0]] = row[1]; });
    const maxSpots  = parseInt(config['max_spots'] || 10);
    const isActive  = config['is_active'] === true || config['is_active'] === 'TRUE' || config['is_active'] === 'true';
    const regData   = regSheet.getDataRange().getValues();
    const confirmed = regData.filter((row, i) => i > 0 && row[4] === 'Confirmed').length;
    return response({ isActive, sessionName: config['session_name'] || 'DYT Session', date: config['date'] || '', time: config['time'] || '', location: config['location'] || 'Roehampton Sport & Fitness Centre', maxSpots, confirmed, spotsLeft: Math.max(0, maxSpots - confirmed) });
  }

  if (action === 'session_register') {
    const name      = e.parameter.name    || '';
    const handle    = ('@' + (e.parameter.handle || '').replace(/^@/, '')).toLowerCase();
    const first     = e.parameter.first === 'true';
    const email     = e.parameter.email  || '';
    const sessionId = e.parameter.id     || null;
    if (!name || !handle || handle === '@') return response({ status: 'error', message: 'Missing name or handle' });

    const noShow = isNoShow(handle);

    const regSheet = getOrCreateSheet('Session Registrations', ['Timestamp', 'Name', 'Handle', 'First Session', 'Status', 'Email', 'Session_ID']);
    if (sessionId) {
      const sessionsSheet = getOrCreateSheet('Sessions', ['ID', 'Name', 'Date', 'Time', 'Location', 'Max Spots', 'Status', 'Drop Date', 'Drop Time']);
      const sessionsData  = sessionsSheet.getDataRange().getValues();
      let sessionRow = null;
      for (let i = 1; i < sessionsData.length; i++) {
        if (String(sessionsData[i][0]).trim() === String(sessionId).trim()) { sessionRow = sessionsData[i]; break; }
      }
      if (!sessionRow) return response({ status: 'error', message: 'Session not found' });
      const status = String(sessionRow[6] || 'soon').toLowerCase().trim();
      if (status !== 'live') return response({ status: 'error', message: 'Session is not open' });
      const maxSpots = parseInt(sessionRow[5] || 10);
      const regData  = regSheet.getDataRange().getValues();
      const duplicate = regData.slice(1).find(row => (row[2] || '').toLowerCase() === handle && String(row[6] || '').trim() === String(sessionId).trim());
      if (duplicate) return response({ status: duplicate[4] === 'Confirmed' ? 'confirmed' : 'waitlist', duplicate: true });
      const confirmed = regData.filter((row, i) => i > 0 && String(row[6]).trim() === String(sessionId).trim() && row[4] === 'Confirmed').length;
      const regStatus = (noShow || confirmed >= maxSpots) ? 'Waitlist' : 'Confirmed';
      regSheet.appendRow([new Date(), name, handle, first ? 'Yes' : 'No', regStatus, email, sessionId]);
      try { if (email) sendConfirmationEmail(email, regStatus, sessionRow[1], sessionRow[2], sessionRow[3], sessionRow[4], noShow, first); } catch(err) { Logger.log(err); }
      return response({ status: regStatus.toLowerCase(), spotsLeft: Math.max(0, maxSpots - confirmed - 1), noShow });
    }
    const configSheet = getOrCreateSheet('Session Config', ['Field', 'Value']);
    const config = {};
    configSheet.getDataRange().getValues().forEach((row, i) => { if (i > 0 && row[0]) config[row[0]] = row[1]; });
    const isActive = config['is_active'] === true || config['is_active'] === 'TRUE' || config['is_active'] === 'true';
    if (!isActive) return response({ status: 'error', message: 'No active session' });
    const maxSpots = parseInt(config['max_spots'] || 10);
    const regData  = regSheet.getDataRange().getValues();
    const duplicate = regData.slice(1).find(row => (row[2] || '').toLowerCase() === handle);
    if (duplicate) return response({ status: duplicate[4] === 'Confirmed' ? 'confirmed' : 'waitlist', duplicate: true });
    const confirmed = regData.filter((row, i) => i > 0 && row[4] === 'Confirmed').length;
    const regStatus = (noShow || confirmed >= maxSpots) ? 'Waitlist' : 'Confirmed';
    regSheet.appendRow([new Date(), name, handle, first ? 'Yes' : 'No', regStatus, email, '']);
    try { if (email) sendConfirmationEmail(email, regStatus, config['session_name'], config['date'], config['time'], config['location'], noShow, first); } catch(err) { Logger.log(err); }
    return response({ status: regStatus.toLowerCase(), spotsLeft: Math.max(0, maxSpots - confirmed - 1), noShow });
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
      name: "DYT Breakfast Club",
      subject: "You're In — " + sName + " · " + sDate,
      body: (first ? "You're in. Welcome to the family." : "You're in. Welcome back, family.") + "\n\nSession: " + sName + "\nDate: " + sDate + "\nTime: " + sTime + "\nLocation: " + sLoc + "\n\nCan't make it? Give 12 hours notice.\nNo-shows sit out the next drop. No exceptions.\n\nFor Hoopers. By Hoopers. DYT Family."
    });
  } else if (noShow) {
    MailApp.sendEmail({
      to: email,
      name: "DYT Breakfast Club",
      subject: "Waitlist — " + sName + " · " + sDate,
      body: "You're on the waitlist for this drop.\n\nYou previously missed a session without giving 12 hours notice.\nAs a result you've been moved to the waitlist for this drop.\n\nIf a spot opens up we'll be in touch.\nAttend this session and you'll be fully reinstated for future drops.\n\nFor Hoopers. By Hoopers. DYT Family."
    });
  } else {
    MailApp.sendEmail({
      to: email,
      name: "DYT Breakfast Club",
      subject: "Waitlist — " + sName + " · " + sDate,
      body: "All 10 spots are taken.\n\nYou're on the waitlist for " + sName + " on " + sDate + " at " + sTime + ".\n\nIf a spot opens up we'll be in touch as soon as possible.\n\nFor Hoopers. By Hoopers. DYT Family."
    });
  }
}

function testEmail() {
  MailApp.sendEmail({ to: 'team@simplybritishballers.com', subject: 'DYT Test Email', body: 'Email is working.', name: 'DYT Breakfast Club' });
}

function response(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
