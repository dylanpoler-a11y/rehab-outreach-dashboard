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

// Field key -> Google Sheet column header mapping
var FIELD_TO_HEADER = {
  'status': 'Status', 'priority': 'Priority', 'type': 'Type',
  'states': 'State(s)', 'keyContact': 'Key Contact',
  'ebitda': 'EBITDA / Financials', 'askingPrice': 'Asking Price',
  'ndaStatus': 'NDA Status', 'dataRoom': 'Data Room', 'siteVisit': 'Site Visit',
  'nextAction': 'Next Action', 'actionOwner': 'Action Owner',
  'deadline': 'Deadline', 'notes': 'Notes', 'lastUpdate': 'Last Update'
};

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var dealName = data.dealName;

    // Open the Pipeline Dashboard tab
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Pipeline Dashboard');

    if (!sheet) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Sheet "Pipeline Dashboard" not found' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    // Find header row (look for "Facility Name" in column A or B)
    var headerRow = -1;
    var nameCol = -1;

    for (var r = 0; r < Math.min(values.length, 10); r++) {
      for (var c = 0; c < values[r].length; c++) {
        var cell = String(values[r][c]).trim();
        if (cell === 'Facility Name') {
          headerRow = r;
          nameCol = c;
        }
      }
      if (headerRow >= 0) break;
    }

    if (headerRow < 0 || nameCol < 0) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Could not find "Facility Name" header' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Build column index map from headers
    var colMap = {};
    for (var c = 0; c < values[headerRow].length; c++) {
      colMap[String(values[headerRow][c]).trim()] = c;
    }

    // Find the deal row by facility name
    var dealRow = -1;
    for (var r = headerRow + 1; r < values.length; r++) {
      var name = String(values[r][nameCol]).trim();
      if (name === dealName) { dealRow = r; break; }
    }

    if (dealRow < 0) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Deal "' + dealName + '" not found in sheet' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Legacy path: status-only update (from drag-and-drop)
    if (data.newStatus && !data.fields) {
      var statusCol = colMap['Status'];
      if (statusCol === undefined) {
        return ContentService.createTextOutput(
          JSON.stringify({ success: false, error: 'Could not find "Status" column' })
        ).setMimeType(ContentService.MimeType.JSON);
      }
      sheet.getRange(dealRow + 1, statusCol + 1).setValue(data.newStatus);
      return ContentService.createTextOutput(
        JSON.stringify({ success: true, deal: dealName, newStatus: data.newStatus })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    // Multi-field update path (from edit popover)
    if (data.fields) {
      var updatedFields = [];
      for (var fieldKey in data.fields) {
        var header = FIELD_TO_HEADER[fieldKey];
        if (!header || colMap[header] === undefined) continue;
        sheet.getRange(dealRow + 1, colMap[header] + 1).setValue(data.fields[fieldKey]);
        updatedFields.push(fieldKey);
      }
      return ContentService.createTextOutput(
        JSON.stringify({ success: true, deal: dealName, updated: updatedFields })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: 'Missing newStatus or fields' })
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
  var sheet = ss.getSheetByName('Pipeline Dashboard');
  if (sheet) {
    Logger.log('Found sheet: ' + sheet.getName() + ' with ' + sheet.getLastRow() + ' rows');
  } else {
    Logger.log('ERROR: Could not find "Pipeline Dashboard" tab');
  }
}

/**
 * RUN THIS ONCE to populate Airtable IDs in your Pipeline Dashboard sheet.
 * Go to Apps Script editor > select "populateAirtableIds" > click Run
 *
 * It will:
 * 1. Find or create an "Airtable ID" column
 * 2. Match each facility name to its Airtable record ID
 * 3. Fill in the IDs
 */
function populateAirtableIds() {
  var DEAL_TO_AIRTABLE_ID = {
    'Nola Detox & Recovery Center': 'recWmGPd7p9F3Alhp',
    'The Grove Recovery': 'recgfSxdxg3XU5jSw',
    'Second Chances': 'recbpyyzLMyXD3OHX',
    'Serenity Grove': 'recjZdOxDxcDpFhem',
    'Williamsville Wellness': 'recfbhn65z7qqk0rB',
    'Serenity Treatment Centers': 'recDKpKlCMRePZFeh',
    'Sanctuary': 'reciYZNJbZOBaivjt',
    'Diamond Recovery': 'recFQ8QnVrYbNBg4j',
    'Vizion Health': 'rec3iBLy02knkSaIz',
    'Clearfork Academy': 'recWoP2HcCtfclrBZ',
    'Asheville Detox (Healthcare Alliance)': 'recPTIvdwSjx1ZL8M',
    'Southeast Detox / Addiction Ctr': 'recH5x5VAaDCjs9kj',
    'Recovery Now / Longbranch': 'recnhD5zQaaiP8fA7',
    'Recovery Bay Center': 'reczoLMGDwfvivZE7',
    'Momentum Recovery': 'recDV6x81cS67WO78',
    'New Waters': 'recWOe4SAg0zsFzV5',
    'DreamLife / Crestview': 'rec1UYfbZr9sVx9tR',
    'Sycamour': 'recDJP3QqxVkZzvqd',
    'Regions Behavioral Hospital': 'reckZg3hi0q40ekNW',
    'Riverwalk Ranch': 'recjmwvYJHmQwSJqX',
    '7 Summit Pathways': 'recN9SqQmoK7jmrNh',
    'Woodlake Center': 'recmhGqK46nOhEd7x',
    'New Vista / Ethan Crossing': 'recMgyNGOxaaH0Js0',
    'New Hope Carolinas': 'recls959IkJ2zvSzA',
    'Advanced Rapid Detox': 'recTXRgsoOTxi3xXG',
    'Cardinal': 'recce3UVJ3OPZG1ZM',
    'The Sylvia Brafman MH Center': 'recvVoTmKEhwGXk7Y',
    'Genesis Behavioral Hospital': 'recLIchoLIbsawX1s',
    'GHR': 'recSfkUTI3ZJ5fLgS',
    'Peachtree Detox (Evoraa)': 'rec0davfHEq47x2hz',
    'Revive Recover': 'recRVboX9ZsZlR2YS',
    'Southern Sky': 'recuuA90QUFSzFgPS',
    'The Wave': 'recU5PgUeHNpgeDTp',
    'Southern Live Oak Wellness': 'reccZf1w6nXo8TyHl',
    'Turning Leaf Behavioral Health': 'recOK29eETN3aFPkZ',
    'Centric': 'recpKzO8XqtpD0TqH',
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Pipeline Dashboard');
  if (!sheet) { Logger.log('ERROR: Sheet "Pipeline Dashboard" not found'); return; }

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  // Find header row
  var headerRow = -1;
  var nameCol = -1;
  for (var r = 0; r < Math.min(values.length, 10); r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (String(values[r][c]).trim() === 'Facility Name') {
        headerRow = r;
        nameCol = c;
        break;
      }
    }
    if (headerRow >= 0) break;
  }

  if (headerRow < 0) { Logger.log('ERROR: Could not find "Facility Name" header'); return; }

  // Find or create "Airtable ID" column
  var aidCol = -1;
  for (var c = 0; c < values[headerRow].length; c++) {
    if (String(values[headerRow][c]).trim() === 'Airtable ID') {
      aidCol = c;
      break;
    }
  }

  if (aidCol < 0) {
    // Add it as the last column
    aidCol = values[headerRow].length;
    sheet.getRange(headerRow + 1, aidCol + 1).setValue('Airtable ID');
    Logger.log('Created "Airtable ID" column at column ' + (aidCol + 1));
  }

  // Fill in IDs
  var filled = 0;
  for (var r = headerRow + 1; r < values.length; r++) {
    var name = String(values[r][nameCol]).trim();
    if (!name) continue;
    var aid = DEAL_TO_AIRTABLE_ID[name];
    if (aid) {
      sheet.getRange(r + 1, aidCol + 1).setValue(aid);
      filled++;
    }
  }

  Logger.log('Done! Filled ' + filled + ' Airtable IDs out of ' + (values.length - headerRow - 1) + ' rows');
}
