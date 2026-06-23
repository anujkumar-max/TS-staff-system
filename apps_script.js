function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const sheetParam = e.parameter.sheet || '';
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Helper to find sheets dynamically
    const getStaffSheet = () => getSheetByPossibleNames(ss, ['Staff_Directory', 'Staff', 'Staff_Master', 'All Staff'], 'staff');
    const getHubSheet = () => getSheetByPossibleNames(ss, ['hub_entries', 'Hub', 'Hub Entries', 'Command Register'], 'hub');
    const getVcSheet = () => getSheetByPossibleNames(ss, ['VC_Schedule', 'VC', 'VC Schedule', 'VC_entries', 'Meetings'], 'vc');
    const getNoticesSheet = () => getSheetByPossibleNames(ss, ['Notices', 'Notice'], 'notice');

    // Specific Action: Verify Login
    if (action === 'login') {
      const phone    = e.parameter.phone || '';
      const password = e.parameter.password || '';
      const sheet = getStaffSheet();
      if (!sheet) return jsonOut({ success: false, error: 'Staff sheet not found' });
      const rows    = sheet.getDataRange().getValues();
      const headers = rows[0];
      const phoneCol = findColIndex(headers, ['Phone', 'Mobile']);
      const passCol  = findColIndex(headers, ['Password']);
      if (phoneCol === -1) return jsonOut({ success: false, error: 'Phone column not found in sheet' });
      
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][phoneCol]).trim() === phone.trim()) {
          const storedPass = passCol >= 0 ? String(rows[i][passCol]).trim() : '1111';
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

    // Specific Action: getStaff
    if (action === 'getStaff') {
      const sheet = getStaffSheet();
      if (!sheet) return jsonOut({ error: 'Staff sheet not found' });
      const rows = sheet.getDataRange().getValues();
      if (rows.length < 2) return jsonOut({ staff: [], data: [] });
      const headers = rows[0];
      const staff = rows.slice(1)
        .filter(r => r[0] !== '' || r[1] !== '')
        .map((row, idx) => {
          const obj = { _rowNumber: idx + 2 };
          headers.forEach((h, i) => obj[h] = row[i] || '');
          return obj;
        });
      return jsonOut({ staff, data: staff });
    }

    // Generic sheet query or specific action fallbacks
    let targetSheetName = '';
    if (sheetParam) {
      targetSheetName = sheetParam;
    } else if (action === 'getHub') {
      targetSheetName = 'hub_entries';
    } else if (action === 'getVC') {
      targetSheetName = 'VC_Schedule';
    }

    if (targetSheetName) {
      let sheet = null;
      const lowerTarget = targetSheetName.toLowerCase();
      if (lowerTarget.indexOf('staff') >= 0) sheet = getStaffSheet();
      else if (lowerTarget.indexOf('hub') >= 0 || lowerTarget.indexOf('entry') >= 0 || lowerTarget.indexOf('command') >= 0) sheet = getHubSheet();
      else if (lowerTarget.indexOf('vc') >= 0 || lowerTarget.indexOf('meeting') >= 0 || lowerTarget.indexOf('conference') >= 0) sheet = getVcSheet();
      else if (lowerTarget.indexOf('notice') >= 0) sheet = getNoticesSheet();
      
      if (!sheet) sheet = ss.getSheetByName(targetSheetName);
      if (!sheet) return jsonOut({ error: 'Sheet not found: ' + targetSheetName });
      
      const rows = sheet.getDataRange().getValues();
      if (rows.length < 1) return jsonOut({ data: [], staff: [], entries: [], vcs: [] });
      
      const headers = rows[0];
      const data = rows.slice(1)
        .filter(r => r[0] !== '' || r[1] !== '')
        .map((row, idx) => {
          const obj = { _rowNumber: idx + 2 };
          headers.forEach((h, i) => obj[h] = row[i] || '');
          return obj;
        });
      return jsonOut({
        data: data,
        staff: data,
        entries: data,
        vcs: data
      });
    }

    return jsonOut({ error: 'Unknown action or missing sheet parameter' });

  } catch(err) {
    return jsonOut({ error: err.message });
  }
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = postData.action || '';
    const sheetName = postData.sheet || '';

    // Helpers to find sheets dynamically
    const getStaffSheet = () => getSheetByPossibleNames(ss, ['Staff_Directory', 'Staff', 'Staff_Master', 'All Staff'], 'staff');
    const getHubSheet = () => getSheetByPossibleNames(ss, ['hub_entries', 'Hub', 'Hub Entries', 'Command Register'], 'hub');
    const getVcSheet = () => getSheetByPossibleNames(ss, ['VC_Schedule', 'VC', 'VC Schedule', 'VC_entries', 'Meetings'], 'vc');
    const getNoticesSheet = () => getSheetByPossibleNames(ss, ['Notices', 'Notice'], 'notice');

    // Specific Action: updateStaff
    if (action === 'updateStaff') {
      const sheet = getStaffSheet();
      if (!sheet) return jsonOut({ error: 'Staff sheet not found' });
      const rows = sheet.getDataRange().getValues();
      const headers = rows[0];
      const phoneCol = findColIndex(headers, ['Phone', 'Mobile']);
      if (phoneCol === -1) return jsonOut({ error: 'Phone column not found in sheet' });

      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][phoneCol]).trim() === String(postData.phone).trim()) {
          const fieldMap = {
            'Name': postData.name, 'Full_Name': postData.name, 'Full Name': postData.name,
            'Rank': postData.rank,
            'Email': postData.email, 'Gmail': postData.email,
            'Department': postData.dept, 'District_Assigned': postData.dept, 'District Assigned': postData.dept,
            'Date of Join': postData.doj, 'Date_Joined': postData.doj, 'Date Joined': postData.doj,
            'Address': postData.address
          };
          Object.entries(fieldMap).forEach(([field, val]) => {
            const col = headers.indexOf(field);
            if (col >= 0 && val !== undefined) sheet.getRange(i+1, col+1).setValue(val);
          });
          return jsonOut({ success: true });
        }
      }
      return jsonOut({ error: 'Staff not found' });
    }

    // Generic database operations by sheet name
    if (sheetName) {
      let sheet = null;
      const lowerSheet = sheetName.toLowerCase();
      if (lowerSheet.indexOf('staff') >= 0) sheet = getStaffSheet();
      else if (lowerSheet.indexOf('hub') >= 0 || lowerSheet.indexOf('entry') >= 0 || lowerSheet.indexOf('command') >= 0) sheet = getHubSheet();
      else if (lowerSheet.indexOf('vc') >= 0 || lowerSheet.indexOf('meeting') >= 0 || lowerSheet.indexOf('conference') >= 0) sheet = getVcSheet();
      else if (lowerSheet.indexOf('notice') >= 0) sheet = getNoticesSheet();
      
      if (!sheet) sheet = ss.getSheetByName(sheetName);
      if (!sheet) return jsonOut({ error: 'Sheet not found: ' + sheetName });

      // Append row
      if (action === 'append' || action === 'add' || action === 'addHub' || action === 'addVC') {
        const rows = sheet.getDataRange().getValues();
        const headers = rows[0];
        let newRow = [];
        if (Array.isArray(postData.data)) {
          newRow = postData.data;
        } else {
          newRow = headers.map(h => postData.data[h] !== undefined ? postData.data[h] : '');
        }
        sheet.appendRow(newRow);
        return jsonOut({ success: true });
      }

      // Update row
      if (action === 'update') {
        const rowNumber = Number(postData.row || postData.rowNumber);
        if (!rowNumber || rowNumber < 2) return jsonOut({ error: 'Invalid row number' });
        const rows = sheet.getDataRange().getValues();
        const headers = rows[0];
        let updatedData = postData.data;
        if (Array.isArray(updatedData)) {
          updatedData.forEach((val, index) => {
            sheet.getRange(rowNumber, index + 1).setValue(val);
          });
        } else {
          Object.entries(updatedData).forEach(([key, val]) => {
            const col = headers.indexOf(key);
            if (col >= 0) sheet.getRange(rowNumber, col + 1).setValue(val);
          });
        }
        return jsonOut({ success: true });
      }

      // Delete row
      if (action === 'delete') {
        const rowNumber = Number(postData.row || postData.rowNumber);
        if (!rowNumber || rowNumber < 2) return jsonOut({ error: 'Invalid row number' });
        sheet.deleteRow(rowNumber);
        return jsonOut({ success: true });
      }
    }

    return jsonOut({ error: 'Unknown action or missing sheet parameter' });

  } catch(err) {
    return jsonOut({ error: err.message });
  }
}

// Helper: case-insensitive match for possible sheet names, fallback to substring match
function getSheetByPossibleNames(ss, possibleNames, fallbackKeyword) {
  for (let name of possibleNames) {
    let sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  const sheets = ss.getSheets();
  for (let sheet of sheets) {
    const sName = sheet.getName().toLowerCase();
    for (let name of possibleNames) {
      if (sName === name.toLowerCase()) return sheet;
    }
  }
  if (fallbackKeyword) {
    for (let sheet of sheets) {
      const sName = sheet.getName().toLowerCase();
      if (sName.indexOf(fallbackKeyword.toLowerCase()) >= 0) return sheet;
    }
  }
  return null;
}

// Helper: case-insensitive column lookup
function findColIndex(headers, possibleNames) {
  for (let name of possibleNames) {
    const idx = headers.indexOf(name);
    if (idx >= 0) return idx;
  }
  for (let i = 0; i < headers.length; i++) {
    const headerLower = String(headers[i]).toLowerCase();
    for (let name of possibleNames) {
      if (headerLower === name.toLowerCase()) return i;
    }
  }
  return -1;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
