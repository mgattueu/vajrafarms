// ═══════════════════════════════════════════════════════════════
// VAJRA FARMS — Google Apps Script v2 (Backend for Staff Portal)
// ═══════════════════════════════════════════════════════════════
//
// SETUP: See SETUP.md for deployment instructions.
// After updating, redeploy: Deploy > Manage deployments > Edit > New version > Deploy
// ═══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    switch (action) {
      case 'login':          return j(handleLogin(data));
      case 'submitOrder':    return j(handleSubmitOrder(data));
      case 'getOrders':      return j(handleGetOrders(data));
      case 'getTodayOrders': return j(handleGetTodayOrders(data));
      case 'updateStatus':   return j(handleUpdateStatus(data));
      case 'getUsers':       return j(handleGetUsers());
      case 'addUser':        return j(handleAddUser(data));
      case 'deleteUser':     return j(handleDeleteUser(data));
      case 'getAreas':       return j(handleGetAreas());
      case 'addArea':        return j(handleAddArea(data));
      case 'deleteArea':     return j(handleDeleteArea(data));
      case 'getDashboard':   return j(handleGetDashboard());
      case 'getRoute':       return j(handleGetRoute(data));
      case 'getSettings':    return j(handleGetSettings());
      case 'updateSettings': return j(handleUpdateSettings(data));
      default:               return j({ success: false, message: 'Unknown action.' });
    }
  } catch (err) {
    return j({ success: false, message: err.toString() });
  }
}

function doGet(e) {
  return j({ success: true, message: 'Vajra Farms API v2 is running.' });
}

// ─── Helpers ─────────────────────────────────────────────

function j(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
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
  return getOrCreateSheet('Users', ['Username', 'Password', 'DisplayName', 'Role', 'Area']);
}

function getOrdersSheet() {
  return getOrCreateSheet('Orders', [
    'Timestamp', 'CustomerName', 'CustomerPhone', 'Area', 'Address',
    'Quantity', 'MilkType', 'DeliveryDate', 'DeliveryTime', 'Frequency',
    'CustomerType', 'TrialDays', 'TrialStartDate', 'Status',
    'StaffName', 'StaffUsername', 'Notes', 'MapsAddress'
  ]);
}

function getAreasSheet() {
  return getOrCreateSheet('Areas', ['AreaName']);
}

function getSettingsSheet() {
  return getOrCreateSheet('Settings', ['Key', 'Value']);
}

function getSetting(key, defaultValue) {
  var sheet = getSettingsSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === key) {
      return data[i][1] ? data[i][1].toString() : defaultValue;
    }
  }
  // Seed default
  sheet.appendRow([key, defaultValue]);
  return defaultValue;
}

function setSetting(key, value) {
  var sheet = getSettingsSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

// ─── Seed Defaults ───────────────────────────────────────

function seedDefaults() {
  // Seed admin user
  var uSheet = getUsersSheet();
  var uData = uSheet.getDataRange().getValues();
  if (uData.length <= 1) {
    uSheet.appendRow(['admin', 'vajra2024', 'Admin', 'admin', '']);
    uSheet.appendRow(['staff', 'staff123', 'Staff Member', 'staff', '']);
    uSheet.appendRow(['del1', 'del1234', 'Delivery 1 - Gandhi Chowk', 'delivery', 'Gandhi Chowk & Centre']);
    uSheet.appendRow(['del2', 'del1234', 'Delivery 2 - Wyra Road', 'delivery', 'Wyra Road & Rotary Nagar']);
    uSheet.appendRow(['del3', 'del1234', 'Delivery 3 - Balaji Nagar', 'delivery', 'Balaji Nagar & Khanapuram']);
    uSheet.appendRow(['del4', 'del1234', 'Delivery 4 - Mamillagudem', 'delivery', 'Mamillagudem & Nehru Nagar']);
    uSheet.appendRow(['del5', 'del1234', 'Delivery 5 - Bypass Road', 'delivery', 'Bypass Road & Dwaraka Nagar']);
    uSheet.appendRow(['del6', 'del1234', 'Delivery 6 - Munneru Side', 'delivery', 'Munneru Side & Raghunadha Palem']);
  }

  // Seed areas
  var aSheet = getAreasSheet();
  var aData = aSheet.getDataRange().getValues();
  if (aData.length <= 1) {
    var defaultAreas = [
      'Gandhi Chowk & Centre',
      'Wyra Road & Rotary Nagar',
      'Balaji Nagar & Khanapuram',
      'Mamillagudem & Nehru Nagar',
      'Bypass Road & Dwaraka Nagar',
      'Munneru Side & Raghunadha Palem'
    ];
    for (var i = 0; i < defaultAreas.length; i++) {
      aSheet.appendRow([defaultAreas[i]]);
    }
  }
}

// ─── Login ───────────────────────────────────────────────

function handleLogin(data) {
  seedDefaults();
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
        role: rows[i][3],
        area: rows[i][4] || ''
      };
    }
  }
  return { success: false, message: 'Invalid username or password.' };
}

// ─── Submit Order ────────────────────────────────────────

function handleSubmitOrder(data) {
  var sheet = getOrdersSheet();
  var trialStart = data.customerType === 'Trial' ? (data.deliveryDate || new Date().toISOString().split('T')[0]) : '';
  sheet.appendRow([
    data.timestamp || new Date().toLocaleString('en-IN', {timeZone: 'Asia/Kolkata'}),
    data.customerName || '',
    data.customerPhone || '',
    data.area || '',
    data.customerAddress || '',
    data.quantity || '',
    data.milkType || 'Buffalo Milk',
    data.deliveryDate || '',
    data.deliveryTime || '',
    data.frequency || '',
    data.customerType || 'Trial',
    data.trialDays || 3,
    trialStart,
    'Active',
    data.staffName || '',
    data.staffUsername || '',
    data.notes || '',
    data.mapsAddress || ''
  ]);
  return { success: true, message: 'Order submitted.' };
}

// ─── Get All Orders ──────────────────────────────────────

function handleGetOrders(data) {
  var sheet = getOrdersSheet();
  var rows = sheet.getDataRange().getValues();
  var orders = [];
  var filterArea = (data && data.area) ? data.area : '';
  var filterStatus = (data && data.status) ? data.status : '';
  var filterType = (data && data.customerType) ? data.customerType : '';

  for (var i = rows.length - 1; i >= 1; i--) {
    var order = {
      row: i + 1,
      timestamp: rows[i][0],
      customerName: rows[i][1],
      customerPhone: rows[i][2],
      area: rows[i][3],
      customerAddress: rows[i][4],
      quantity: rows[i][5],
      milkType: rows[i][6],
      deliveryDate: rows[i][7],
      deliveryTime: rows[i][8],
      frequency: rows[i][9],
      customerType: rows[i][10],
      trialDays: rows[i][11],
      trialStartDate: rows[i][12],
      status: rows[i][13],
      staffName: rows[i][14],
      notes: rows[i][16],
      mapsAddress: rows[i][17] || ''
    };
    if (filterArea && order.area !== filterArea) continue;
    if (filterStatus && order.status !== filterStatus) continue;
    if (filterType && order.customerType !== filterType) continue;
    orders.push(order);
  }
  return { success: true, orders: orders };
}

// ─── Get Today's Orders ──────────────────────────────────

function handleGetTodayOrders(data) {
  var sheet = getOrdersSheet();
  var rows = sheet.getDataRange().getValues();
  var today = new Date();
  var todayStr = Utilities.formatDate(today, 'Asia/Kolkata', 'yyyy-MM-dd');
  var orders = [];
  var filterArea = (data && data.area) ? data.area : '';

  for (var i = 1; i < rows.length; i++) {
    var delivDate = rows[i][7] ? rows[i][7].toString() : '';
    if (delivDate instanceof Date) {
      delivDate = Utilities.formatDate(delivDate, 'Asia/Kolkata', 'yyyy-MM-dd');
    }
    var status = rows[i][13] ? rows[i][13].toString() : '';
    if (status !== 'Active') continue;

    // Check if delivery date matches today
    // Also include recurring orders (Daily, Alternate Days)
    var freq = rows[i][9] ? rows[i][9].toString() : '';
    var isToday = (delivDate === todayStr);
    var isRecurring = (freq === 'Daily');

    // For alternate days, check if the difference from start is even
    if (freq === 'Alternate Days' && delivDate) {
      var startDate = new Date(delivDate);
      var diffDays = Math.floor((today - startDate) / (1000 * 60 * 60 * 24));
      if (diffDays >= 0 && diffDays % 2 === 0) isRecurring = true;
    }

    // For weekly, check same day of week
    if (freq === 'Weekly' && delivDate) {
      var startDay = new Date(delivDate).getDay();
      if (today.getDay() === startDay) isRecurring = true;
    }

    // Check trial expiry
    var custType = rows[i][10] ? rows[i][10].toString() : '';
    var trialDays = rows[i][11] ? parseInt(rows[i][11]) : 3;
    var trialStart = rows[i][12] ? rows[i][12].toString() : '';
    if (custType === 'Trial' && trialStart) {
      var trialStartDate = new Date(trialStart);
      var daysSinceStart = Math.floor((today - trialStartDate) / (1000 * 60 * 60 * 24));
      if (daysSinceStart >= trialDays) continue; // Trial expired
    }

    if (!isToday && !isRecurring) continue;

    var order = {
      row: i + 1,
      timestamp: rows[i][0],
      customerName: rows[i][1],
      customerPhone: rows[i][2],
      area: rows[i][3],
      customerAddress: rows[i][4],
      quantity: rows[i][5],
      milkType: rows[i][6],
      deliveryDate: delivDate,
      deliveryTime: rows[i][8],
      frequency: freq,
      customerType: custType,
      trialDays: trialDays,
      trialStartDate: trialStart,
      status: status,
      staffName: rows[i][14],
      notes: rows[i][16],
      mapsAddress: rows[i][17] || ''
    };
    if (filterArea && order.area !== filterArea) continue;
    orders.push(order);
  }

  return { success: true, orders: orders };
}

// ─── Update Customer Status ──────────────────────────────

function handleUpdateStatus(data) {
  var sheet = getOrdersSheet();
  var row = data.row;
  if (!row) return { success: false, message: 'Row not specified.' };

  if (data.status) sheet.getRange(row, 14).setValue(data.status);
  if (data.customerType) sheet.getRange(row, 11).setValue(data.customerType);
  if (data.trialDays) sheet.getRange(row, 12).setValue(data.trialDays);

  return { success: true, message: 'Updated.' };
}

// ─── Dashboard Analytics ─────────────────────────────────

function handleGetDashboard() {
  var sheet = getOrdersSheet();
  var rows = sheet.getDataRange().getValues();
  var today = new Date();
  var todayStr = Utilities.formatDate(today, 'Asia/Kolkata', 'yyyy-MM-dd');

  var trialActive = 0, kathaActive = 0;
  var trialLitres = 0, kathaLitres = 0;
  var trialInquiries = 0, kathaInquiries = 0;
  var paymentLapsed = 0, inactive = 0;
  var todayTrialOrders = 0, todayKathaOrders = 0;
  var todayTrialLitres = 0, todayKathaLitres = 0;

  // Last 7 days data for charts
  var dailyData = {};
  for (var d = 6; d >= 0; d--) {
    var date = new Date(today);
    date.setDate(date.getDate() - d);
    var dateStr = Utilities.formatDate(date, 'Asia/Kolkata', 'yyyy-MM-dd');
    dailyData[dateStr] = { trial: 0, katha: 0, trialLitres: 0, kathaLitres: 0 };
  }

  for (var i = 1; i < rows.length; i++) {
    var custType = rows[i][10] ? rows[i][10].toString() : 'Trial';
    var status = rows[i][13] ? rows[i][13].toString() : 'Active';
    var qty = parseFloat(rows[i][5]) || 0;
    var delivDate = rows[i][7] ? rows[i][7].toString() : '';
    if (delivDate instanceof Date) {
      delivDate = Utilities.formatDate(delivDate, 'Asia/Kolkata', 'yyyy-MM-dd');
    }

    // Count by type and status
    if (custType === 'Trial') {
      trialInquiries++;
      if (status === 'Active') { trialActive++; trialLitres += qty; }
    } else {
      kathaInquiries++;
      if (status === 'Active') { kathaActive++; kathaLitres += qty; }
    }
    if (status === 'Payment Lapsed') paymentLapsed++;
    if (status === 'Inactive') inactive++;

    // Today's data
    if (delivDate === todayStr && status === 'Active') {
      if (custType === 'Trial') { todayTrialOrders++; todayTrialLitres += qty; }
      else { todayKathaOrders++; todayKathaLitres += qty; }
    }

    // Daily chart data
    if (dailyData[delivDate] !== undefined) {
      if (custType === 'Trial') {
        dailyData[delivDate].trial++;
        dailyData[delivDate].trialLitres += qty;
      } else {
        dailyData[delivDate].katha++;
        dailyData[delivDate].kathaLitres += qty;
      }
    }
  }

  var chartLabels = Object.keys(dailyData);
  var chartTrial = [], chartKatha = [], chartTrialL = [], chartKathaL = [];
  for (var k = 0; k < chartLabels.length; k++) {
    var key = chartLabels[k];
    chartTrial.push(dailyData[key].trial);
    chartKatha.push(dailyData[key].katha);
    chartTrialL.push(dailyData[key].trialLitres);
    chartKathaL.push(dailyData[key].kathaLitres);
  }

  return {
    success: true,
    summary: {
      trialActive: trialActive,
      kathaActive: kathaActive,
      trialLitres: trialLitres,
      kathaLitres: kathaLitres,
      trialInquiries: trialInquiries,
      kathaInquiries: kathaInquiries,
      paymentLapsed: paymentLapsed,
      inactive: inactive,
      todayTrialOrders: todayTrialOrders,
      todayKathaOrders: todayKathaOrders,
      todayTrialLitres: todayTrialLitres,
      todayKathaLitres: todayKathaLitres
    },
    chart: {
      labels: chartLabels,
      trialCount: chartTrial,
      kathaCount: chartKatha,
      trialLitres: chartTrialL,
      kathaLitres: chartKathaL
    }
  };
}

// ─── Areas ───────────────────────────────────────────────

function handleGetAreas() {
  seedDefaults();
  var sheet = getAreasSheet();
  var data = sheet.getDataRange().getValues();
  var areas = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) areas.push(data[i][0].toString());
  }
  return { success: true, areas: areas };
}

function handleAddArea(data) {
  var name = (data.areaName || '').trim();
  if (!name) return { success: false, message: 'Area name required.' };
  var sheet = getAreasSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase() === name.toLowerCase()) {
      return { success: false, message: 'Area already exists.' };
    }
  }
  sheet.appendRow([name]);
  return { success: true, message: 'Area added.' };
}

function handleDeleteArea(data) {
  var name = (data.areaName || '').trim();
  var sheet = getAreasSheet();
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString() === name) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Area removed.' };
    }
  }
  return { success: false, message: 'Area not found.' };
}

// ─── Users ───────────────────────────────────────────────

function handleGetUsers() {
  seedDefaults();
  var sheet = getUsersSheet();
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    users.push({
      username: data[i][0],
      displayName: data[i][2],
      role: data[i][3],
      area: data[i][4] || ''
    });
  }
  return { success: true, users: users };
}

function handleAddUser(data) {
  var sheet = getUsersSheet();
  var rows = sheet.getDataRange().getValues();
  var username = (data.username || '').toLowerCase().trim();
  if (!username || !data.password || !data.displayName) {
    return { success: false, message: 'All fields are required.' };
  }
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase().trim() === username) {
      return { success: false, message: 'Username already exists.' };
    }
  }
  sheet.appendRow([username, data.password, data.displayName, data.role || 'staff', data.area || '']);
  return { success: true, message: 'User added.' };
}

function handleDeleteUser(data) {
  var sheet = getUsersSheet();
  var rows = sheet.getDataRange().getValues();
  var username = (data.username || '').toLowerCase().trim();
  if (username === 'admin') return { success: false, message: 'Cannot delete admin.' };
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0].toString().toLowerCase().trim() === username) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'User removed.' };
    }
  }
  return { success: false, message: 'User not found.' };
}

// ─── Route Optimization ────────────────────────────────

// Haversine formula: distance in km between two lat/lng points
function haversine(lat1, lng1, lat2, lng2) {
  var R = 6371; // Earth radius in km
  var dLat = (lat2 - lat1) * Math.PI / 180;
  var dLng = (lng2 - lng1) * Math.PI / 180;
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
          Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
          Math.sin(dLng / 2) * Math.sin(dLng / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}

function handleGetRoute(data) {
  var filterArea = (data && data.area) ? data.area : '';
  var defaultOrigin = '4-2-744/2, Mamatha Hospital Rd, Dwaraka Nagar, Kaviraj Nagar, Khammam, Telangana 507002, India';
  var origin = (data && data.origin) ? data.origin : getSetting('routeOrigin', defaultOrigin);

  // Get today's active orders (reuse logic from getTodayOrders)
  var result = handleGetTodayOrders({ area: filterArea });
  if (!result.success || !result.orders || result.orders.length === 0) {
    return { success: false, message: 'No deliveries found for today.' };
  }

  var orders = result.orders;

  // Build full addresses — prefer MapsAddress column, fallback to regular address
  var stops = [];
  for (var i = 0; i < orders.length; i++) {
    var mapsAddr = orders[i].mapsAddress || '';
    var addr = orders[i].customerAddress || '';
    var area = orders[i].area || '';

    // Use Google Maps address if available, otherwise build from regular address
    var fullAddr = '';
    if (mapsAddr) {
      fullAddr = mapsAddr;
    } else {
      fullAddr = addr;
      if (area && fullAddr.toLowerCase().indexOf(area.toLowerCase()) === -1) {
        fullAddr += ', ' + area;
      }
      if (fullAddr.toLowerCase().indexOf('khammam') === -1) {
        fullAddr += ', Khammam, Telangana';
      }
    }
    stops.push({
      index: i,
      address: fullAddr,
      name: orders[i].customerName || '',
      phone: orders[i].customerPhone || '',
      quantity: orders[i].quantity || '',
      milkType: orders[i].milkType || '',
      time: orders[i].deliveryTime || ''
    });
  }

  // Geocode all addresses and optimize using nearest-neighbor
  var optimizedOrder = [];
  var totalDistance = '';
  var totalDuration = '';

  try {
    // Geocode origin
    var originGeo = Maps.newGeocoder().geocode(origin);
    var originLat = 17.2473; // Khammam default
    var originLng = 80.1514;
    if (originGeo.status === 'OK' && originGeo.results.length > 0) {
      originLat = originGeo.results[0].geometry.location.lat;
      originLng = originGeo.results[0].geometry.location.lng;
    }

    // Geocode each stop
    for (var g = 0; g < stops.length; g++) {
      try {
        var geo = Maps.newGeocoder().geocode(stops[g].address);
        if (geo.status === 'OK' && geo.results.length > 0) {
          stops[g].lat = geo.results[0].geometry.location.lat;
          stops[g].lng = geo.results[0].geometry.location.lng;
        } else {
          // Can't geocode — assign default coords so it goes last
          stops[g].lat = originLat;
          stops[g].lng = originLng;
        }
      } catch (ge) {
        stops[g].lat = originLat;
        stops[g].lng = originLng;
      }
    }

    // Nearest-neighbor algorithm
    if (stops.length >= 2) {
      var visited = [];
      var unvisited = stops.slice();
      var curLat = originLat;
      var curLng = originLng;
      var totalDist = 0;

      while (unvisited.length > 0) {
        var nearest = 0;
        var nearestDist = haversine(curLat, curLng, unvisited[0].lat, unvisited[0].lng);
        for (var n = 1; n < unvisited.length; n++) {
          var d = haversine(curLat, curLng, unvisited[n].lat, unvisited[n].lng);
          if (d < nearestDist) { nearest = n; nearestDist = d; }
        }
        totalDist += nearestDist;
        curLat = unvisited[nearest].lat;
        curLng = unvisited[nearest].lng;
        visited.push(unvisited[nearest]);
        unvisited.splice(nearest, 1);
      }

      optimizedOrder = visited;
      totalDistance = totalDist.toFixed(1) + ' km';
      // Rough estimate: 20 km/h average in town
      totalDuration = Math.round((totalDist / 20) * 60) + ' min';
    } else if (stops.length === 1) {
      optimizedOrder = stops;
      var d1 = haversine(originLat, originLng, stops[0].lat, stops[0].lng);
      totalDistance = d1.toFixed(1) + ' km';
      totalDuration = Math.round((d1 / 20) * 60) + ' min';
    }
  } catch (e) {
    // If geocoding fails entirely, return original order
    optimizedOrder = stops.length > 0 ? stops : [];
  }

  // Build Google Maps navigation URLs — auto-split into chunks of 10
  var routes = [];
  if (optimizedOrder.length > 0) {
    var chunkSize = 10;
    var totalChunks = Math.ceil(optimizedOrder.length / chunkSize);
    for (var c = 0; c < totalChunks; c++) {
      var start = c * chunkSize;
      var end = Math.min(start + chunkSize, optimizedOrder.length);
      var chunk = optimizedOrder.slice(start, end);

      // Origin: for first chunk use the farm origin, for subsequent chunks use last stop of previous chunk
      var routeOrigin = (c === 0) ? origin : optimizedOrder[start - 1].address;
      var routeDest = chunk[chunk.length - 1].address;

      var wpStr = '';
      if (chunk.length > 1) {
        var wpAddrs = [];
        for (var p = 0; p < chunk.length - 1; p++) {
          wpAddrs.push(encodeURIComponent(chunk[p].address));
        }
        wpStr = '&waypoints=' + wpAddrs.join('|');
      }

      var url = 'https://www.google.com/maps/dir/?api=1'
        + '&origin=' + encodeURIComponent(routeOrigin)
        + '&destination=' + encodeURIComponent(routeDest)
        + wpStr
        + '&travelmode=driving';

      routes.push({
        routeNumber: c + 1,
        totalRoutes: totalChunks,
        stops: chunk,
        stopRange: (start + 1) + '-' + end,
        mapsUrl: url
      });
    }
  }

  return {
    success: true,
    stops: optimizedOrder,
    totalStops: optimizedOrder.length,
    totalDistance: totalDistance,
    totalDuration: totalDuration,
    routes: routes,
    mapsUrl: routes.length > 0 ? routes[0].mapsUrl : ''
  };
}

// ─── Settings ───────────────────────────────────────────

function handleGetSettings() {
  var defaultOrigin = '4-2-744/2, Mamatha Hospital Rd, Dwaraka Nagar, Kaviraj Nagar, Khammam, Telangana 507002, India';
  return {
    success: true,
    settings: {
      routeOrigin: getSetting('routeOrigin', defaultOrigin)
    }
  };
}

function handleUpdateSettings(data) {
  if (data.routeOrigin !== undefined) {
    setSetting('routeOrigin', data.routeOrigin);
  }
  return { success: true, message: 'Settings updated.' };
}
