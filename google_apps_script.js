/**
 * Google Apps Script — Pipeline Status Writeback
 *
 * HOW TO SET UP:
 * 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/19w2nLn7VrNEWQSRaEIgPQuZwPi88ABNKWnl8Bh7pqVk
 * 2. Go to Extensions > Apps Script
 * 3. Delete any existing code and paste this entire file
 * 4. Click "Deploy" > "New deployment"
 * 5. Click the gear icon next to "Select type" and choose "Web app"
 * 6. Set:
 *    - Description: "Pipeline Status Updater"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 * 7. Click "Deploy"
 * 8. Authorize the script when prompted
 * 9. Copy the Web App URL (looks like: https://script.google.com/macros/s/xxxxx/exec)
 * 10. Paste that URL into your index.html where it says APPS_SCRIPT_URL = ''
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var dealName = data.dealName;
    var newStatus = data.newStatus;

    if (!dealName || !newStatus) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Missing dealName or newStatus' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Open the Pipeline Detail tab
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Pipeline Detail');

    if (!sheet) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Sheet "Pipeline Detail" not found' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    // Find header row (look for "Facility Name" in column A or B)
    var headerRow = -1;
    var nameCol = -1;
    var statusCol = -1;

    for (var r = 0; r < Math.min(values.length, 10); r++) {
      for (var c = 0; c < values[r].length; c++) {
        var cell = String(values[r][c]).trim();
        if (cell === 'Facility Name') {
          headerRow = r;
          nameCol = c;
        }
        if (cell === 'Status' && headerRow === r) {
          statusCol = c;
        }
      }
      if (headerRow >= 0) break;
    }

    if (headerRow < 0 || nameCol < 0) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Could not find "Facility Name" header' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (statusCol < 0) {
      // If Status wasn't on the same row scan, look for it explicitly
      for (var c = 0; c < values[headerRow].length; c++) {
        if (String(values[headerRow][c]).trim() === 'Status') {
          statusCol = c;
          break;
        }
      }
    }

    if (statusCol < 0) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Could not find "Status" column' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Find the deal row by facility name
    var updated = false;
    for (var r = headerRow + 1; r < values.length; r++) {
      var name = String(values[r][nameCol]).trim();
      if (name === dealName) {
        // Update the status cell (rows and cols are 1-indexed in setCell)
        sheet.getRange(r + 1, statusCol + 1).setValue(newStatus);
        updated = true;
        break;
      }
    }

    if (!updated) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Deal "' + dealName + '" not found in sheet' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ success: true, deal: dealName, newStatus: newStatus })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// Test function — run this from Apps Script to verify the sheet is accessible
function testAccess() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Pipeline Detail');
  if (sheet) {
    Logger.log('Found sheet: ' + sheet.getName() + ' with ' + sheet.getLastRow() + ' rows');
  } else {
    Logger.log('ERROR: Could not find "Pipeline Detail" tab');
  }
}
