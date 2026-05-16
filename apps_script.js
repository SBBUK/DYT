// ── DYT Apps Script ──
// Paste this into Extensions → Apps Script in your Google Sheet
// Then: Deploy → New Deployment → Web App → Execute as: Me → Who has access: Anyone → Deploy
// Copy the web app URL and paste it into session.html where it says YOUR_APPS_SCRIPT_URL_HERE

// ─────────────────────────────────────────
// GOOGLE SHEET SETUP
// ─────────────────────────────────────────
// You need a sheet named "Sessions" with these columns (row 1 = headers):
//   ID | Name | Date | Time | Location | Max Spots | Status
//
// Example rows:
//   1 | DYT Breakfast Club   | 19/05 | 7:00 AM | Roehampton Sport & Fitness Centre | 10 | live
//   2 | DYT Evening Session  | 22/05 | 6:00 PM | Roehampton Sport & Fitness Centre | 10 | soon
//
// Status values:
//   live  → Drop is open. Claim Your Spot button is active.
//   soon  → Drop is coming. Card shown with "Drop Coming Soon" label, no button.
//   closed → Session done. Card shown dimmed at the bottom.
//
// The "Session Registrations" sheet columns should be:
//   Timestamp | Name | Handle | First Session | Status | Email | Session_ID
// ─────────────────────────────────────────

const SHEET_ID = '1-Yb9u6h0v6BQffZqOhV7kjYXPtSjqqPWC0cyzgKxWCY';

// ─────────────────────────────────────────
// MILESTONE HELPERS
// ─────────────────────────────────────────
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

// ─────────────────────────────────────────
// GET HANDLER
// ─────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';

  // ── Multi-session: return all sessions with confirmed spot counts ──
  if (action === 'sessions_list') {
    const sessionsSheet = getOrCreateSheet('Sessions', ['ID', 'Name', 'Date', 'Time', 'Location', 'Max Spots', 'Status']);
    const regSheet      = getOrCreateSheet('Session Registrations', ['Timestamp', 'Name', 'Handle', 'First Session', 'Status', 'Email', 'Session_ID']);

    const sessionsData = sessionsSheet.getDataRange().getValues();
    const regData      = regSheet.getDataRange().getValues();

    // Build a map of sessionId -> confirmed count
    const confirmedMap = {};
    for (let i = 1; i < regData.length; i++) {
      const row = regData[i];
      if (!row[0]) continue;
      // Status is column index 4, Session_ID is column index 6
      const sid    = String(row[6] || '').trim();
      const status = String(row[4] || '').trim();
      if (sid && status === 'Confirmed') {
        confirmedMap[sid] = (confirmedMap[sid] || 0) + 1;
      }
    }

    const sessions = [];
    for (let i = 1; i < sessionsData.length; i++) {
      const row = sessionsData[i];
      if (!row[0]) continue;
      const id        = String(row[0]).trim();
      const maxSpots  = parseInt(row[5] || 10);
      const confirmed = confirmedMap[id] || 0;
      sessions.push({
        id,
        name:      row[1] || '',
        date:      row[2] || '',
        time:      row[3] || '',
        location:  row[4] || '',
        maxSpots,
        confirmed,
        spotsLeft: Math.max(0, maxSpots - confirmed),
        status:    String(row[6] || 'soon').toLowerCase().trim(),
        dropDate:  row[7] || '',
        dropTime:  row[8] || ''
      });
    }

    return response(sessions);
  }

  // ── Multi-session: return info for a specific session by ID ──
  if (action === 'session_info') {
    const sessionId = e.parameter.id || null;

    // If an ID is provided, look it up in the Sessions sheet
    if (sessionId) {
      const sessionsSheet = getOrCreateSheet('Sessions', ['ID', 'Name', 'Date', 'Time', 'Location', 'Max Spots', 'Status']);
      const regSheet      = getOrCreateSheet('Session Registrations', ['Timestamp', 'Name', 'Handle', 'First Session', 'Status', 'Email', 'Session_ID']);

      const sessionsData = sessionsSheet.getDataRange().getValues();
      let sessionRow = null;
      for (let i = 1; i < sessionsData.length; i++) {
        if (String(sessionsData[i][0]).trim() === String(sessionId).trim()) {
          sessionRow = sessionsData[i];
          break;
        }
      }

      if (!sessionRow) return response({ error: 'Session not found', isActive: false });

      const maxSpots  = parseInt(sessionRow[5] || 10);
      const status    = String(sessionRow[6] || 'soon').toLowerCase().trim();
      const isActive  = status === 'live';

      const regData   = regSheet.getDataRange().getValues();
      const confirmed = regData.filter((row, i) => i > 0 && String(row[6]).trim() === String(sessionId).trim() && row[4] === 'Confirmed').length;

      return response({
        isActive,
        id:          String(sessionRow[0]).trim(),
        sessionName: sessionRow[1] || 'DYT Session',
        date:        sessionRow[2] || '',
        time:        sessionRow[3] || '',
        location:    sessionRow[4] || 'Roehampton Sport & Fitness Centre',
        maxSpots,
        confirmed,
        spotsLeft:   Math.max(0, maxSpots - confirmed),
        status
      });
    }

    // No ID — fall back to legacy Session Config sheet behaviour
    const configSheet = getOrCreateSheet('Session Config', ['Field', 'Value']);
    const regSheet    = getOrCreateSheet('Session Registrations', ['Timestamp', 'Name', 'Handle', 'First Session', 'Status', 'Email', 'Session_ID']);

    const config = {};
    const configData = configSheet.getDataRange().getValues();
    for (let i = 1; i < configData.length; i++) {
      if (configData[i][0]) config[configData[i][0]] = configData[i][1];
    }

    const maxSpots  = parseInt(config['max_spots'] || 10);
    const isActive  = config['is_active'] === true || config['is_active'] === 'TRUE' || config['is_active'] === 'true';
    const regData   = regSheet.getDataRange().getValues();
    const confirmed = regData.filter((row, i) => i > 0 && row[4] === 'Confirmed').length;

    return response({
      isActive,
      sessionName: config['session_name'] || 'DYT Session',
      date:        config['date']         || '',
      time:        config['time']         || '',
      location:    config['location']     || 'Roehampton Sport & Fitness Centre',
      maxSpots,
      confirmed,
      spotsLeft:   Math.max(0, maxSpots - confirmed)
    });
  }

  // ── Session Drop: register for a specific session (id-aware) ──
  if (action === 'session_register') {
    const name      = e.parameter.name    || '';
    const handle    = ('@' + (e.parameter.handle || '').replace(/^@/, '')).toLowerCase();
    const first     = e.parameter.first === 'true';
    const email     = e.parameter.email  || '';
    const sessionId = e.parameter.id     || null;

    if (!name || !handle || handle === '@') {
      return response({ status: 'error', message: 'Missing name or handle' });
    }

    const regSheet = getOrCreateSheet('Session Registrations', ['Timestamp', 'Name', 'Handle', 'First Session', 'Status', 'Email', 'Session_ID']);

    // ── ID-aware path: look up session from Sessions sheet ──
    if (sessionId) {
      const sessionsSheet = getOrCreateSheet('Sessions', ['ID', 'Name', 'Date', 'Time', 'Location', 'Max Spots', 'Status']);
      const sessionsData  = sessionsSheet.getDataRange().getValues();
      let sessionRow = null;
      for (let i = 1; i < sessionsData.length; i++) {
        if (String(sessionsData[i][0]).trim() === String(sessionId).trim()) {
          sessionRow = sessionsData[i];
          break;
        }
      }

      if (!sessionRow) return response({ status: 'error', message: 'Session not found' });

      const status = String(sessionRow[6] || 'soon').toLowerCase().trim();
      if (status !== 'live') return response({ status: 'error', message: 'Session is not open for registration' });

      const maxSpots = parseInt(sessionRow[5] || 10);
      const regData  = regSheet.getDataRange().getValues();

      // Check for duplicate (same handle + same session)
      const duplicate = regData.slice(1).find(row =>
        (row[2] || '').toLowerCase() === handle &&
        String(row[6] || '').trim() === String(sessionId).trim()
      );
      if (duplicate) {
        const dupStatus = duplicate[4];
        return response({ status: dupStatus === 'Confirmed' ? 'confirmed' : 'waitlist', duplicate: true });
      }

      const confirmed  = regData.filter((row, i) => i > 0 && String(row[6]).trim() === String(sessionId).trim() && row[4] === 'Confirmed').length;
      const regStatus  = confirmed < maxSpots ? 'Confirmed' : 'Waitlist';

      regSheet.appendRow([new Date(), name, handle, first ? 'Yes' : 'No', regStatus, email, sessionId]);

      // Send confirmation email
      const sName = sessionRow[1] || 'DYT Session';
      const sDate = sessionRow[2] || '';
      const sTime = sessionRow[3] || '';
      const sLoc  = sessionRow[4] || 'Roehampton Sport & Fitness Centre';
      try { if (email) { sendConfirmationEmail(email, regStatus, sName, sDate, sTime, sLoc); } } catch(mailErr) { Logger.log('Email error: ' + mailErr); }

      return response({ status: regStatus.toLowerCase(), spotsLeft: Math.max(0, maxSpots - confirmed - 1) });
    }

    // ── Legacy path: use Session Config sheet ──
    const configSheet = getOrCreateSheet('Session Config', ['Field', 'Value']);
    const config = {};
    configSheet.getDataRange().getValues().forEach((row, i) => {
      if (i > 0 && row[0]) config[row[0]] = row[1];
    });

    const isActive = config['is_active'] === true || config['is_active'] === 'TRUE' || config['is_active'] === 'true';
    if (!isActive) return response({ status: 'error', message: 'No active session' });

    const maxSpots = parseInt(config['max_spots'] || 10);
    const regData  = regSheet.getDataRange().getValues();

    const duplicate = regData.slice(1).find(row => (row[2] || '').toLowerCase() === handle);
    if (duplicate) {
      const dupStatus = duplicate[4];
      return response({ status: dupStatus === 'Confirmed' ? 'confirmed' : 'waitlist', duplicate: true });
    }

    const confirmed = regData.filter((row, i) => i > 0 && row[4] === 'Confirmed').length;
    const regStatus = confirmed < maxSpots ? 'Confirmed' : 'Waitlist';

    regSheet.appendRow([new Date(), name, handle, first ? 'Yes' : 'No', regStatus, email, '']);

    const sName = config['session_name'] || 'DYT Session';
    const sDate = config['date']         || '';
    const sTime = config['time']         || '';
    const sLoc  = config['location']     || 'Roehampton Sport & Fitness Centre';
    try { if (email) { sendConfirmationEmail(email, regStatus, sName, sDate, sTime, sLoc); } } catch(mailErr) { Logger.log('Email error: ' + mailErr); }

    return response({ status: regStatus.toLowerCase(), spotsLeft: Math.max(0, maxSpots - confirmed - 1) });
  }

  // ── Tracker: return all members ──
  if (action === 'all') {
    const sheet = getMembersSheet();
    const data  = sheet.getDataRange().getValues();
    const members = [];
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      members.push({
        name:      data[i][0],
        handle:    data[i][1],
        sessions:  data[i][2] || 0,
        milestone: data[i][4] || ''
      });
    }
    return response(members);
  }

  // ── Tracker: look up a single member ──
  if (action === 'lookup') {
    const handle = (e.parameter.handle || '').toLowerCase().replace('@', '');
    const sheet  = getMembersSheet();
    const data   = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      if ((row[1] || '').toLowerCase().replace('@', '') === handle) {
        const sessions = row[2] || 0;
        return response({
          found:     true,
          name:      row[0],
          handle:    row[1],
          sessions,
          milestone: row[4] || '',
          next:      getNextMilestone(sessions)
        });
      }
    }
    return response({ found: false });
  }

  return response({ error: 'Unknown action' });
}

// ─────────────────────────────────────────
// POST HANDLER (tracker check-in + register)
// ─────────────────────────────────────────
function doPost(e) {
  const body  = JSON.parse(e.postData.contents);
  const sheet = getMembersSheet();
  const data  = sheet.getDataRange().getValues();

  // ── Check in a list of handles ──
  if (body.action === 'checkin') {
    const handles  = (body.handles || []).map(h => h.toLowerCase().replace('@', ''));
    const achieved = [];
    for (let i = 1; i < data.length; i++) {
      const rowHandle = (data[i][1] || '').toLowerCase().replace('@', '');
      if (!rowHandle || !handles.includes(rowHandle)) continue;
      const oldCount    = data[i][2] || 0;
      const newCount    = oldCount + 1;
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

  // ── Register a new member to the tracker ──
  if (body.action === 'register') {
    sheet.appendRow([body.name, body.handle, 0, new Date(), '']);
    return response({ success: true });
  }

  return response({ error: 'Unknown action' });
}

// ─────────────────────────────────────────
// EMAIL HELPER
// ─────────────────────────────────────────
function sendConfirmationEmail(email, status, sName, sDate, sTime, sLoc) {
  if (status === 'Confirmed') {
    MailApp.sendEmail({
      to: email,
      subject: "You're In. — DYT Session Drop",
      htmlBody: `<div style="background:#080808;color:#ffffff;font-family:monospace;padding:48px;max-width:560px;margin:0 auto;">
        <p style="font-size:11px;letter-spacing:0.4em;color:#FF4400;text-transform:uppercase;margin-bottom:24px;">Session Confirmed</p>
        <h1 style="font-family:sans-serif;font-size:48px;line-height:1;font-weight:900;margin-bottom:32px;">You're In.<br>See You On<br>The Court.</h1>
        <div style="border-top:1px solid rgba(255,255,255,0.15);border-bottom:1px solid rgba(255,255,255,0.15);padding:24px 0;margin-bottom:32px;">
          <p style="font-size:11px;letter-spacing:0.3em;color:#FF4400;text-transform:uppercase;margin-bottom:4px;">Session</p>
          <p style="font-size:20px;font-weight:700;margin-bottom:20px;">${sName}</p>
          <p style="font-size:11px;letter-spacing:0.3em;color:#FF4400;text-transform:uppercase;margin-bottom:4px;">Date & Time</p>
          <p style="font-size:20px;font-weight:700;margin-bottom:20px;">${sDate} · ${sTime}</p>
          <p style="font-size:11px;letter-spacing:0.3em;color:#FF4400;text-transform:uppercase;margin-bottom:4px;">Location</p>
          <p style="font-size:20px;font-weight:700;">${sLoc}</p>
        </div>
        <p style="font-size:12px;letter-spacing:0.15em;color:rgba(255,255,255,0.5);line-height:2;margin-bottom:32px;">Can't make it? Give <strong style="color:#ffffff;">12 hours notice</strong>. No-shows sit out the next drop. No exceptions.</p>
        <p style="font-size:10px;letter-spacing:0.35em;color:rgba(255,255,255,0.25);text-transform:uppercase;">For Hoopers. By Hoopers. DYT Family.</p>
      </div>`
    });
  } else {
    MailApp.sendEmail({
      to: email,
      subject: "You're on the Waitlist — DYT Session Drop",
      htmlBody: `<div style="background:#080808;color:#ffffff;font-family:monospace;padding:48px;max-width:560px;margin:0 auto;">
        <p style="font-size:11px;letter-spacing:0.4em;color:#FF4400;text-transform:uppercase;margin-bottom:24px;">Waitlist</p>
        <h1 style="font-family:sans-serif;font-size:48px;line-height:1;font-weight:900;margin-bottom:32px;">All 10 Spots<br>Are Taken.</h1>
        <p style="font-size:13px;letter-spacing:0.12em;color:rgba(255,255,255,0.6);line-height:2;margin-bottom:32px;">You're on the waitlist for ${sName} on ${sDate} at ${sTime}.<br><br>If a spot opens up we'll be in touch as soon as possible. Keep an eye on your inbox.</p>
        <p style="font-size:10px;letter-spacing:0.35em;color:rgba(255,255,255,0.25);text-transform:uppercase;">For Hoopers. By Hoopers. DYT Family.</p>
      </div>`
    });
  }
}

function response(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
