// ═══════════════════════════════════════════════════════════════
// VAJRA FARMS — Google Apps Script (Backend for Staff Portal)
// ═══════════════════════════════════════════════════════════════
//
// SETUP INSTRUCTIONS:
// 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/15muhKJlmGZXL7cN5zhHvosonlU6Dx_nVbNEHD3RkCMc/edit
// 2. Go to Extensions > Apps Script
// 3. Delete any existing code and paste this entire file
// 4. Click "Deploy" > "New deployment"
// 5. Choose type: "Web app"
// 6. Set "Execute as": "Me"
// 7. Set "Who has access": "Anyone"
// 8. Click "Deploy" and authorize when prompted
// 9. Copy the Web App URL
// 10. Paste that URL into portal.html where it says YOUR_GOOGLE_APPS_SCRIPT_URL_HERE
//
// SHEET STRUCTURE (will be auto-created on first run):
// Sheet "Orders" — all milk orders
// Sheet "Users"  — staff/admin login credentials
// ═══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;

    switch (action) {
      case 'login':
        return jsonResponse(handleLogin(data));
      case 'submitOrder':
        return jsonResponse(handleSubmitOrder(data));
      case 'getOrders':
        return jsonResponse(handleGetOrders());
      case 'getUsers':
        return jsonResponse(handleGetUsers());
      case 'addUser':
        return jsonResponse(handleAddUser(data));
      case 'deleteUser':
        return jsonResponse(handleDeleteUser(data));
      default:
        return jsonResponse({ success: false, message: 'Unknown action.' });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function doGet(e) {
  return jsonResponse({ success: true, message: 'Vajra Farms API is running.' });
}

// ─────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function getUsersSheet() {
  return getOrCreateSheet('Users', ['Username', 'Password', 'DisplayName', 'Role']);
}

function getOrdersSheet() {
  return getOrCreateSheet('Orders', [
    'Timestamp', 'CustomerName', 'CustomerPhone', 'CustomerAddress',
    'Quantity (L)', 'MilkType', 'DeliveryDate', 'DeliveryTime',
    'Frequency', 'StaffName', 'StaffUsername', 'Notes'
  ]);
}

// ─────────────────────────────────────────────
// SEED DEFAULT ADMIN (runs once if Users sheet is empty)
// ─────────────────────────────────────────────

function seedDefaultAdmin() {
  var sheet = getUsersSheet();
  var data = sheet.getDataRange().getValues();
  // If only header row exists, add default admin
  if (data.length <= 1) {
    sheet.appendRow(['admin', 'vajra2024', 'Admin', 'admin']);
    sheet.appendRow(['staff', 'staff123', 'Staff Member', 'staff']);
  }
}

// ─────────────────────────────────────────────
// LOGIN
// ─────────────────────────────────────────────

function handleLogin(data) {
  seedDefaultAdmin();
  var sheet = getUsersSheet();
  var rows = sheet.getDataRange().getValues();
  var username = (data.username || '').toLowerCase().trim();
  var password = data.password || '';

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase().trim() === username &&
        rows[i][1].toString() === password) {
      return {
        success: true,
        username: rows[i][0],
        displayName: rows[i][2],
        role: rows[i][3]
      };
    }
  }

  return { success: false, message: 'Invalid username or password.' };
}

// ─────────────────────────────────────────────
// SUBMIT ORDER
// ─────────────────────────────────────────────

function handleSubmitOrder(data) {
  var sheet = getOrdersSheet();
  sheet.appendRow([
    data.timestamp || new Date().toLocaleString(),
    data.customerName || '',
    data.customerPhone || '',
    data.customerAddress || '',
    data.quantity || '',
    data.milkType || '',
    data.deliveryDate || '',
    data.deliveryTime || '',
    data.frequency || '',
    data.staffName || '',
    data.staffUsername || '',
    data.notes || ''
  ]);
  return { success: true, message: 'Order submitted.' };
}

// ─────────────────────────────────────────────
// GET ORDERS (Admin)
// ─────────────────────────────────────────────

function handleGetOrders() {
  var sheet = getOrdersSheet();
  var data = sheet.getDataRange().getValues();
  var orders = [];

  for (var i = data.length - 1; i >= 1; i--) {
    orders.push({
      timestamp: data[i][0],
      customerName: data[i][1],
      customerPhone: data[i][2],
      customerAddress: data[i][3],
      quantity: data[i][4],
      milkType: data[i][5],
      deliveryDate: data[i][6],
      deliveryTime: data[i][7],
      frequency: data[i][8],
      staffName: data[i][9],
      notes: data[i][11]
    });
  }

  return { success: true, orders: orders };
}

// ─────────────────────────────────────────────
// GET USERS (Admin)
// ─────────────────────────────────────────────

function handleGetUsers() {
  seedDefaultAdmin();
  var sheet = getUsersSheet();
  var data = sheet.getDataRange().getValues();
  var users = [];

  for (var i = 1; i < data.length; i++) {
    users.push({
      username: data[i][0],
      displayName: data[i][2],
      role: data[i][3]
    });
  }

  return { success: true, users: users };
}

// ─────────────────────────────────────────────
// ADD USER (Admin)
// ─────────────────────────────────────────────

function handleAddUser(data) {
  var sheet = getUsersSheet();
  var rows = sheet.getDataRange().getValues();
  var username = (data.username || '').toLowerCase().trim();

  if (!username || !data.password || !data.displayName) {
    return { success: false, message: 'All fields are required.' };
  }

  // Check for duplicate username
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase().trim() === username) {
      return { success: false, message: 'Username already exists.' };
    }
  }

  sheet.appendRow([username, data.password, data.displayName, data.role || 'staff']);
  return { success: true, message: 'User added.' };
}

// ─────────────────────────────────────────────
// DELETE USER (Admin)
// ─────────────────────────────────────────────

function handleDeleteUser(data) {
  var sheet = getUsersSheet();
  var rows = sheet.getDataRange().getValues();
  var username = (data.username || '').toLowerCase().trim();

  if (username === 'admin') {
    return { success: false, message: 'Cannot delete the admin account.' };
  }

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase().trim() === username) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'User removed.' };
    }
  }

  return { success: false, message: 'User not found.' };
}
