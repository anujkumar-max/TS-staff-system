function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const sheetParam = e.parameter.sheet || '';
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Specific Action: Verify Login
    if (action === 'login') {
      const phone    = e.parameter.phone || '';
      const password = e.parameter.password || '';
      const sheet = ss.getSheetByName('Staff') || ss.getSheetByName('Staff_Master');
      if (!sheet) return jsonOut({ success: false, error: 'Staff sheet not found' });
      const rows    = sheet.getDataRange().getValues();
      const headers = rows[0];
      const phoneCol = headers.indexOf('Phone');
      const passCol  = headers.indexOf('Password');
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

    // Specific Action: getStaff
    if (action === 'getStaff') {
      const sheet = ss.getSheetByName('Staff') || ss.getSheetByName('Staff_Master');
      if (!sheet) return jsonOut({ error: 'Staff sheet not found' });
      const rows = sheet.getDataRange().getValues();
      if (rows.length < 2) return jsonOut({ staff: [], data: [] });
      const headers = rows[0];
      const staff = rows.slice(1)
        .filter(r => r[0] !== '')
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
      let sheet = ss.getSheetByName(targetSheetName);
      if (!sheet) {
        if (targetSheetName === 'Staff' || targetSheetName === 'Staff_Master') {
          sheet = ss.getSheetByName('Staff') || ss.getSheetByName('Staff_Master');
        } else if (targetSheetName === 'Hub' || targetSheetName === 'hub_entries') {
          sheet = ss.getSheetByName('Hub') || ss.getSheetByName('hub_entries');
        } else if (targetSheetName === 'VC' || targetSheetName === 'VC_Schedule') {
          sheet = ss.getSheetByName('VC') || ss.getSheetByName('VC_Schedule');
        }
      }
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

    // Specific Action: updateStaff
    if (action === 'updateStaff') {
      const sheet = ss.getSheetByName('Staff') || ss.getSheetByName('Staff_Master');
      if (!sheet) return jsonOut({ error: 'Staff sheet not found' });
      const rows = sheet.getDataRange().getValues();
      const headers = rows[0];
      const phoneCol = headers.indexOf('Phone');
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
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        if (sheetName === 'Staff' || sheetName === 'Staff_Master') {
          sheet = ss.getSheetByName('Staff') || ss.getSheetByName('Staff_Master');
        } else if (sheetName === 'Hub' || sheetName === 'hub_entries') {
          sheet = ss.getSheetByName('Hub') || ss.getSheetByName('hub_entries');
        } else if (sheetName === 'VC' || sheetName === 'VC_Schedule') {
          sheet = ss.getSheetByName('VC') || ss.getSheetByName('VC_Schedule');
        }
      }
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

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
