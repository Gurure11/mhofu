/**
 * ============================================================
 * MineLink — Google Apps Script Backend
 * ============================================================
 * Paste this entire file into your Apps Script editor
 * (Extensions > Apps Script in your Google Sheet).
 *
 * SETUP STEPS:
 * 1. Deploy > New deployment > Web app.
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 2. Copy the Web App URL and paste it into the HTML file's API_URL.
 *
 * The script will AUTO-CREATE the required sheet tabs and headers
 * on the first request. No manual sheet setup needed.
 *
 * IMPORTANT: The Sheet ID is sent by the HTML frontend with every
 * request, so this backend can work with any sheet without hardcoding.
 * ============================================================
 */

// ========== DEFAULT SHEET ID (fallback) ==========
const DEFAULT_SHEET_ID = '1RlRRh_7sbJJAWWPvlRkKb0iaCuIQZWdixpTVYKLUYFQ';

// ========== SHEET CONFIGURATION ==========
const SHEETS = {
  MINES: {
    name: 'Mines',
    headers: [
      'id',
      'name',
      'location',
      'price',
      'coords',
      'phone',
      'email',
      'whatsapp',
      'images',
      'description',
      'seller',
      'timestamp',
      'isActive'
    ]
  },
  USERS: {
    name: 'Users',
    headers: [
      'id',
      'email',
      'password',
      'fullName',
      'phone',
      'createdAt',
      'lastLogin'
    ]
  }
};

// ========== ENTRY POINTS ==========

/**
 * Handles GET requests.
 * The sheetId can be passed as a query parameter; falls back to DEFAULT_SHEET_ID.
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = String(params.action || '').trim();
    const sheetId = String(params.sheetId || DEFAULT_SHEET_ID).trim();
    const ss = SpreadsheetApp.openById(sheetId);
    ensureAllSheets(ss);

    switch (action) {
      case 'getMines':
        return jsonResponse(getMines(ss));
      case 'getUser':
        return jsonResponse(getUserByEmail(ss, params.email));
      case 'health':
        return jsonResponse({ success: true, message: 'MineLink API is running.', sheetId: sheetId });
      default:
        return jsonResponse({ error: 'Invalid GET action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message, stack: err.stack });
  }
}

/**
 * Handles POST requests.
 * The sheetId can be passed in the JSON body; falls back to DEFAULT_SHEET_ID.
 */
function doPost(e) {
  try {
    let payload;
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      payload = (e && e.parameter) ? e.parameter : {};
    }

    const sheetId = String(payload.sheetId || DEFAULT_SHEET_ID).trim();
    const ss = SpreadsheetApp.openById(sheetId);
    ensureAllSheets(ss);

    const action = String(payload.action || '').trim();

    switch (action) {
      case 'addMine':
        return jsonResponse(addMine(ss, payload.mine || {}));
      case 'login':
        return jsonResponse(loginUser(ss, payload.email, payload.password));
      case 'register':
        return jsonResponse(registerUser(ss, payload));
      case 'updateMine':
        return jsonResponse(updateMine(ss, payload.mine || {}));
      case 'deleteMine':
        return jsonResponse(deleteMine(ss, payload.id));
      default:
        return jsonResponse({ error: 'Invalid POST action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ error: err.message, stack: err.stack });
  }
}

// ========== SHEET MANAGEMENT ==========

/**
 * Ensures all required sheet tabs exist and have correct headers.
 * Creates them if missing. Also adds a demo user if Users is empty.
 */
function ensureAllSheets(ss) {
  Object.values(SHEETS).forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
    }
    ensureHeaders(sheet, cfg.headers);
  });

  // Seed a demo user if Users sheet is empty
  const userSheet = ss.getSheetByName(SHEETS.USERS.name);
  const userData = userSheet.getDataRange().getValues();
  if (userData.length <= 1) {
    const demo = [
      'user-demo-001',
      'demo@minelink.com',
      hashPassword('demo123'),
      'Demo Miner',
      '+264811234567',
      new Date().toISOString(),
      ''
    ];
    userSheet.appendRow(demo);
  }
}

/**
 * Ensures the given sheet has the exact headers in row 1.
 */
function ensureHeaders(sheet, expectedHeaders) {
  const lastRow = sheet.getLastRow();
  let currentHeaders = [];

  if (lastRow >= 1) {
    const headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1);
    currentHeaders = headerRange.getValues()[0];
  }

  const current = currentHeaders.map(h => String(h).trim());
  const expected = expectedHeaders.map(h => String(h).trim());
  const matches = expected.every((h, i) => current[i] === h) && current.length === expected.length;

  if (!matches) {
    sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), expected.length) + 5).clearContent();
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    sheet.getRange(1, 1, 1, expected.length)
      .setFontWeight('bold')
      .setBackground('#0a2e4e')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
}

// ========== MINES LOGIC ==========

function getMines(ss) {
  const sheet = ss.getSheetByName(SHEETS.MINES.name);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return { success: true, mines: [] };
  }

  const headers = data[0].map(h => String(h).trim());
  const mines = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (!row[headers.indexOf('name')]) continue;

    const mine = {};
    headers.forEach((header, idx) => {
      let val = row[idx];
      if (header === 'images' && typeof val === 'string') {
        val = val.split('\n').map(s => s.trim()).filter(Boolean);
      }
      if (header === 'isActive') {
        val = (val === true || String(val).toLowerCase() === 'true' || val === 1);
      }
      mine[header] = val;
    });

    if (mine.isActive === false) continue;
    mines.push(mine);
  }

  mines.sort((a, b) => new Date(b.timestamp || 0) - new Date(a.timestamp || 0));
  return { success: true, mines };
}

function addMine(ss, mine) {
  const sheet = ss.getSheetByName(SHEETS.MINES.name);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  const rowData = headers.map(header => {
    switch (header) {
      case 'id':
        return 'mine-' + Utilities.getUuid();
      case 'images':
        return Array.isArray(mine.images) ? mine.images.join('\n') : (mine.images || '');
      case 'timestamp':
        return new Date().toISOString();
      case 'isActive':
        return true;
      case 'seller':
        return mine.seller || 'unknown';
      case 'price':
        return mine.price !== undefined && mine.price !== null && mine.price !== '' ? Number(mine.price) : '';
      default:
        return mine[header] !== undefined && mine[header] !== null ? mine[header] : '';
    }
  });

  sheet.appendRow(rowData);
  return { success: true, id: rowData[headers.indexOf('id')] };
}

function updateMine(ss, mine) {
  if (!mine.id) return { error: 'Missing mine ID.' };
  const sheet = ss.getSheetByName(SHEETS.MINES.name);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idCol = headers.indexOf('id');
  if (idCol === -1) return { error: 'ID column not found.' };

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idCol]).trim() === String(mine.id).trim()) {
      const rowRange = sheet.getRange(r + 1, 1, 1, headers.length);
      const currentRow = data[r];
      const newRow = headers.map((header, i) => {
        if (mine[header] !== undefined) {
          if (header === 'images') {
            return Array.isArray(mine.images) ? mine.images.join('\n') : mine.images;
          }
          return mine[header];
        }
        return currentRow[i];
      });
      rowRange.setValues([newRow]);
      return { success: true, id: mine.id };
    }
  }
  return { error: 'Mine not found.' };
}

function deleteMine(ss, id) {
  if (!id) return { error: 'Missing mine ID.' };
  const sheet = ss.getSheetByName(SHEETS.MINES.name);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const idCol = headers.indexOf('id');
  const activeCol = headers.indexOf('isActive');
  if (idCol === -1 || activeCol === -1) return { error: 'Required columns not found.' };

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][idCol]).trim() === String(id).trim()) {
      sheet.getRange(r + 1, activeCol + 1).setValue(false);
      return { success: true };
    }
  }
  return { error: 'Mine not found.' };
}

// ========== USER / AUTH LOGIC ==========

function registerUser(ss, payload) {
  const email = String(payload.email || '').trim().toLowerCase();
  const password = String(payload.password || '');
  if (!email || !password) return { error: 'Email and password are required.' };
  if (password.length < 6) return { error: 'Password must be at least 6 characters.' };

  const sheet = ss.getSheetByName(SHEETS.USERS.name);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const emailCol = headers.indexOf('email');

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][emailCol]).trim().toLowerCase() === email) {
      return { error: 'An account with this email already exists.' };
    }
  }

  const newRow = headers.map(header => {
    switch (header) {
      case 'id':
        return 'user-' + Utilities.getUuid();
      case 'email':
        return email;
      case 'password':
        return hashPassword(password);
      case 'fullName':
        return payload.fullName || '';
      case 'phone':
        return payload.phone || '';
      case 'createdAt':
        return new Date().toISOString();
      case 'lastLogin':
        return '';
      default:
        return '';
    }
  });

  sheet.appendRow(newRow);

  return {
    success: true,
    user: {
      id: newRow[headers.indexOf('id')],
      email: email,
      fullName: payload.fullName || '',
      phone: payload.phone || ''
    }
  };
}

function loginUser(ss, email, password) {
  email = String(email || '').trim().toLowerCase();
  password = String(password || '');
  if (!email || !password) return { error: 'Email and password are required.' };

  const sheet = ss.getSheetByName(SHEETS.USERS.name);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const emailCol = headers.indexOf('email');
  const passCol = headers.indexOf('password');
  const idCol = headers.indexOf('id');
  const nameCol = headers.indexOf('fullName');
  const phoneCol = headers.indexOf('phone');
  const lastLoginCol = headers.indexOf('lastLogin');

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][emailCol]).trim().toLowerCase() === email) {
      const storedHash = String(data[r][passCol] || '');
      if (storedHash === hashPassword(password)) {
        if (lastLoginCol !== -1) {
          sheet.getRange(r + 1, lastLoginCol + 1).setValue(new Date().toISOString());
        }
        return {
          success: true,
          user: {
            id: data[r][idCol] || '',
            email: email,
            fullName: nameCol !== -1 ? data[r][nameCol] : '',
            phone: phoneCol !== -1 ? data[r][phoneCol] : ''
          }
        };
      } else {
        return { error: 'Incorrect password.' };
      }
    }
  }

  return { error: 'No account was found for this email. Please sign up first.' };
}

function getUserByEmail(ss, email) {
  email = String(email || '').trim().toLowerCase();
  if (!email) return { error: 'Email is required.' };

  const sheet = ss.getSheetByName(SHEETS.USERS.name);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const emailCol = headers.indexOf('email');
  const idCol = headers.indexOf('id');
  const nameCol = headers.indexOf('fullName');
  const phoneCol = headers.indexOf('phone');

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][emailCol]).trim().toLowerCase() === email) {
      return {
        success: true,
        user: {
          id: data[r][idCol] || '',
          email: email,
          fullName: nameCol !== -1 ? data[r][nameCol] : '',
          phone: phoneCol !== -1 ? data[r][phoneCol] : ''
        }
      };
    }
  }
  return { error: 'User not found.' };
}

// ========== UTILITIES ==========

function hashPassword(password) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password),
    Utilities.Charset.UTF_8
  );
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ========== OPTIONAL: ONE-TIME SETUP FUNCTION ==========
function setupSheets() {
  const ss = SpreadsheetApp.openById(DEFAULT_SHEET_ID);
  ensureAllSheets(ss);
  Logger.log('Sheets ready: ' + Object.values(SHEETS).map(s => s.name).join(', '));
}