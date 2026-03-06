// ============================================================
// VCO Car Plan - Google Apps Script Web App
//
// SETUP:
// 1. In your Google Sheet, go to Extensions > Apps Script
// 2. Delete any existing code and paste this entire file
// 3. Click Deploy > New Deployment
// 4. Select "Web app" as the type
// 5. Set "Execute as" to "Me"
// 6. Set "Who has access" to "Anyone"
// 7. Click Deploy and authorize when prompted
// 8. Copy the Web App URL and paste it into the dashboard Settings
// ============================================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = data.action;

    if (action === 'saveAssignments') {
      return saveAssignments(ss, data);
    } else if (action === 'updateAvailability') {
      return updateAvailability(ss, data);
    } else if (action === 'addSignout') {
      return addSignout(ss, data);
    } else if (action === 'returnCar') {
      return returnCarLog(ss, data);
    } else if (action === 'updateRoster') {
      return updateRoster(ss, data);
    } else if (action === 'deleteSignout') {
      return deleteSignout(ss, data);
    }

    return respond({ success: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok', message: 'VCO Car Plan API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Save car assignments to Daily Schedule tab
function saveAssignments(ss, data) {
  var sheet = ss.getSheetByName('Daily Schedule');
  if (!sheet) return respond({ success: false, error: 'Daily Schedule tab not found' });

  var assignments = data.assignments || [];
  var date = data.date || '';

  // Set date
  sheet.getRange('B2').setValue(date);

  // Clear existing schedule data (rows 6-21, columns C-M for car assignments)
  sheet.getRange('C6:M21').clearContent();

  // Clear passenger manifest (rows 26+)
  var lastRow = sheet.getLastRow();
  if (lastRow >= 26) {
    sheet.getRange('A26:M' + Math.max(26, lastRow)).clearContent();
  }

  // Write assignments as morning push manifest starting at row 26
  var manifestRow = 26;
  for (var i = 0; i < assignments.length; i++) {
    var a = assignments[i];
    var paxList = a.passengers.join(', ');
    sheet.getRange(manifestRow + i, 1).setValue(a.departTime + 'L');
    sheet.getRange(manifestRow + i, 2).setValue(a.hotel);
    // Find car column (3 = first car column)
    var carCol = findCarColumn(sheet, a.car);
    if (carCol > 0) {
      sheet.getRange(manifestRow + i, carCol).setValue(
        'D: ' + a.driver + (paxList ? '\n' + paxList : '')
      );
    }
  }

  // Also update Availability tab with committed cars
  var availSheet = ss.getSheetByName('Availability');
  if (availSheet) {
    for (var i = 0; i < assignments.length; i++) {
      var a = assignments[i];
      // Find the car row in availability (rows 5-15)
      for (var row = 5; row <= 17; row++) {
        var name = availSheet.getRange(row, 2).getValue();
        if (name === a.car) {
          availSheet.getRange(row, 3).setValue('Committed');
          availSheet.getRange(row, 4).setValue(a.driver);
          availSheet.getRange(row, 5).setValue(a.destination);
          availSheet.getRange(row, 7).setValue(a.departTime + 'L - Morning Push');
          break;
        }
      }
    }
  }

  // Save raw JSON to B3 for dashboard to read back
  sheet.getRange('B3').setValue(JSON.stringify(assignments));

  return respond({ success: true, saved: assignments.length });
}

function findCarColumn(sheet, carName) {
  var headers = sheet.getRange(5, 1, 1, 13).getValues()[0];
  for (var c = 0; c < headers.length; c++) {
    if (headers[c] === carName) return c + 1;
  }
  return -1;
}

// Update car availability status
function updateAvailability(ss, data) {
  var sheet = ss.getSheetByName('Availability');
  if (!sheet) return respond({ success: false, error: 'Availability tab not found' });

  var carName = data.carName;
  var status = data.status;
  var driver = data.driver || '';
  var dest = data.destination || '';
  var eta = data.eta || '';

  for (var row = 5; row <= 17; row++) {
    var name = sheet.getRange(row, 2).getValue();
    if (name === carName) {
      sheet.getRange(row, 3).setValue(status);
      sheet.getRange(row, 4).setValue(driver);
      sheet.getRange(row, 5).setValue(dest);
      sheet.getRange(row, 6).setValue(eta);
      break;
    }
  }

  return respond({ success: true });
}

// Add sign-out log entry
function addSignout(ss, data) {
  var sheet = ss.getSheetByName('Availability');
  if (!sheet) return respond({ success: false, error: 'Availability tab not found' });

  // Find next empty row in sign-out log (starts at row 21)
  var logStart = 21;
  var row = logStart;
  while (sheet.getRange(row, 1).getValue() !== '') {
    row++;
    if (row > 50) break;
  }

  sheet.getRange(row, 1).setNumberFormat('@').setValue(data.time);
  sheet.getRange(row, 2).setValue(data.car);
  sheet.getRange(row, 3).setValue(data.driver);
  sheet.getRange(row, 4).setValue(data.passengers || '');
  sheet.getRange(row, 5).setValue(data.destination);
  sheet.getRange(row, 6).setValue(data.eta);

  // Update car status to committed
  for (var r = 5; r <= 17; r++) {
    var name = sheet.getRange(r, 2).getValue();
    if (name === data.car) {
      sheet.getRange(r, 3).setValue('Committed');
      sheet.getRange(r, 4).setValue(data.driver);
      sheet.getRange(r, 5).setValue(data.destination);
      sheet.getRange(r, 6).setValue(data.eta);
      break;
    }
  }

  return respond({ success: true });
}

// Mark car as returned in sign-out log
function returnCarLog(ss, data) {
  var sheet = ss.getSheetByName('Availability');
  if (!sheet) return respond({ success: false, error: 'Availability tab not found' });

  // Find the sign-out entry by time and car name
  var logStart = 21;
  for (var row = logStart; row <= 50; row++) {
    var time = sheet.getRange(row, 1).getValue().toString().trim();
    var car = sheet.getRange(row, 2).getValue().toString().trim();
    var dataTime = data.time.toString().trim();
    if ((time === dataTime || time === String(parseInt(dataTime))) && car === data.car && sheet.getRange(row, 7).getValue() === '') {
      sheet.getRange(row, 7).setValue(data.timeIn);
      break;
    }
  }

  // Update car status to available
  for (var r = 5; r <= 17; r++) {
    var name = sheet.getRange(r, 2).getValue();
    if (name === data.car) {
      sheet.getRange(r, 3).setValue('Available');
      sheet.getRange(r, 4).setValue('');
      sheet.getRange(r, 5).setValue('');
      sheet.getRange(r, 6).setValue('');
      break;
    }
  }

  return respond({ success: true });
}

// Update roster presence
function updateRoster(ss, data) {
  var sheet = ss.getSheetByName('Roster');
  if (!sheet) return respond({ success: false, error: 'Roster tab not found' });

  var present = data.present || []; // array of names that are present

  // Update column D (Present) for each person
  for (var row = 5; row <= 60; row++) {
    var name = sheet.getRange(row, 1).getValue();
    if (!name) continue;
    var isPresent = present.indexOf(name) >= 0;
    sheet.getRange(row, 4).setValue(isPresent ? 'Yes' : 'No');
  }

  return respond({ success: true });
}

// Delete a sign-out log entry
function deleteSignout(ss, data) {
  var sheet = ss.getSheetByName('Availability');
  if (!sheet) return respond({ success: false, error: 'Availability tab not found' });

  var logStart = 21;
  for (var row = logStart; row <= 50; row++) {
    var time = sheet.getRange(row, 1).getValue().toString().trim();
    var car = sheet.getRange(row, 2).getValue().toString().trim();
    var dataTime = data.time.toString().trim();
    // Handle Sheets converting "0730" to number 730
    if ((time === dataTime || time === String(parseInt(dataTime))) && car === data.car) {
      // Clear the row
      sheet.getRange(row, 1, 1, 7).clearContent();
      // Update car status to available if it was committed
      for (var r = 5; r <= 17; r++) {
        var name = sheet.getRange(r, 2).getValue();
        if (name === data.car) {
          sheet.getRange(r, 3).setValue('Available');
          sheet.getRange(r, 4).setValue('');
          sheet.getRange(r, 5).setValue('');
          sheet.getRange(r, 6).setValue('');
          break;
        }
      }
      break;
    }
  }

  return respond({ success: true });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
