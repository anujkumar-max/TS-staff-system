const STAFF_SHEET = 'Staff';
const HUB_SHEET   = 'Hub';
const VC_SHEET    = 'VC';

function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- GET STAFF LIST ---
    if (action === 'getStaff') {
      const sheet = ss.getSheetByName(STAFF_SHEET);
      if (!sheet) return jsonOut({ error: 'Staff sheet not found' });
      const rows = sheet.getDataRange().getValues();
      if (rows.length < 2) return jsonOut({ staff: [] });
      const headers = rows[0];
      const staff = rows.slice(1)
        .filter(r => r[0] !== '')
        .map(row => {
          const obj = {};
          headers.forEach((h, i) => obj[h] = row[i] || '');
          return obj;
        });
      return jsonOut({ staff });
    }

    // --- GET HUB ENTRIES ---
    if (action === 'getHub') {
      const sheet = ss.getSheetByName(HUB_SHEET);
      if (!sheet) return jsonOut({ error: 'Hub sheet not found' });
      const rows = sheet.getDataRange().getValues();
      if (rows.length < 2) return jsonOut({ entries: [] });
      const headers = rows[0];
      const entries = rows.slice(1)
        .filter(r => r[0] !== '')
        .map(row => {
          const obj = {};
          headers.forEach((h, i) => obj[h] = row[i] || '');
          return obj;
        });
      return jsonOut({ entries });
    }

    // --- GET VC/MEETINGS ---
    if (action === 'getVC') {
      const sheet = ss.getSheetByName(VC_SHEET);
      if (!sheet) return jsonOut({ error: 'VC sheet not found' });
      const rows = sheet.getDataRange().getValues();
      if (rows.length < 2) return jsonOut({ vcs: [] });
      const headers = rows[0];
      const vcs = rows.slice(1)
        .filter(r => r[0] !== '')
        .map(row => {
          const obj = {};
          headers.forEach((h, i) => obj[h] = row[i] || '');
          return obj;
        });
      return jsonOut({ vcs });
    }

    // --- VERIFY LOGIN ---
    if (action === 'login') {
      const phone    = e.parameter.phone || '';
      const password = e.parameter.password || '';
      const sheet    = ss.getSheetByName(STAFF_SHEET);
      if (!sheet) return jsonOut({ success: false, error: 'Staff sheet not found' });
      const rows    = sheet.getDataRange().getValues();
      const headers = rows[0];
      const phoneCol = headers.indexOf('Phone');
      const passCol  = headers.indexOf('Password');
      const roleCol  = headers.indexOf('Role');
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][phoneCol]).trim() === phone.trim()) {
          const storedPass = String(rows[i][passCol]).trim() || '1111';
          if (storedPass === password.trim()) {
            const obj = {};
            headers.forEach((h, j) => obj[h] = rows[i][j] || '');
            return jsonOut({ success: true, user: obj });
          } else {
            return jsonOut({ success: false, error: 'Wrong password' });
          }
        }
      }
      return jsonOut({ success: false, error: 'Phone number not found' });
    }

    return jsonOut({ error: 'Unknown action: ' + action });

  } catch(err) {
    return jsonOut({ error: err.message });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.getActiveSpreadsheet();

    // --- ADD HUB ENTRY ---
    if (data.action === 'addHub') {
      const sheet = ss.getSheetByName(HUB_SHEET);
      if (!sheet) return jsonOut({ error: 'Hub sheet not found' });
      const rows = sheet.getDataRange().getValues();
      const headers = rows[0];
      const newRow = headers.map(h => data[h] || data[h.toLowerCase()] || '');
      newRow[headers.indexOf('Added On') >= 0 ? headers.indexOf('Added On') : 0]
        = new Date().toLocaleString('en-IN');
      sheet.appendRow(newRow);
      return jsonOut({ success: true });
    }

    // --- UPDATE STAFF (Profile Edit) ---
    if (data.action === 'updateStaff') {
      const sheet   = ss.getSheetByName(STAFF_SHEET);
      if (!sheet) return jsonOut({ error: 'Staff sheet not found' });
      const rows    = sheet.getDataRange().getValues();
      const headers = rows[0];
      const phoneCol = headers.indexOf('Phone');
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][phoneCol]).trim() === String(data.phone).trim()) {
          const fieldMap = {
            'Name': data.name, 'Rank': data.rank,
            'Email': data.email, 'Department': data.dept,
            'Date of Join': data.doj, 'Address': data.address
          };
          Object.entries(fieldMap).forEach(([field, val]) => {
            const col = headers.indexOf(field);
            if (col >= 0 && val) sheet.getRange(i+1, col+1).setValue(val);
          });
          return jsonOut({ success: true });
        }
      }
      return jsonOut({ error: 'Staff not found' });
    }

    // --- ADD VC ENTRY ---
    if (data.action === 'addVC') {
      const sheet = ss.getSheetByName(VC_SHEET);
      if (!sheet) return jsonOut({ error: 'VC sheet not found' });
      const rows    = sheet.getDataRange().getValues();
      const headers = rows[0];
      const newRow  = headers.map(h => data[h] || data[h.toLowerCase()] || '');
      sheet.appendRow(newRow);
      return jsonOut({ success: true });
    }

    return jsonOut({ error: 'Unknown action' });

  } catch(err) {
    return jsonOut({ error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
