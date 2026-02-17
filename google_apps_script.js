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

// Field key -> Google Sheet column header mapping (Pipeline Dashboard tab)
var FIELD_TO_HEADER = {
  'status': 'Status', 'priority': 'Priority', 'type': 'Type',
  'states': 'State(s)', 'keyContact': 'Key Contact',
  'ebitda': 'EBITDA / Financials', 'askingPrice': 'Asking Price',
  'ndaStatus': 'NDA Status', 'dataRoom': 'Data Room', 'siteVisit': 'Site Visit',
  'nextAction': 'Next Action', 'actionOwner': 'Action Owner',
  'deadline': 'Deadline', 'notes': 'Notes', 'lastUpdate': 'Last Update'
};

// Action item field key -> Action Items tab column header mapping
var ACTION_FIELD_TO_HEADER = {
  'priority': 'Priority', 'action': 'Action Item', 'facility': 'Facility',
  'owner': 'Owner', 'deadline': 'Deadline', 'status': 'Status',
  'notes': 'Notes', 'pipelineStatus': 'Pipeline Status'
};

// Name of the Action Items tab in Google Sheets
var ACTION_ITEMS_TAB = 'Action Items';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ============================================================
    // ACTION ITEM: Edit existing action item
    // Payload: { actionItem: "text", actionFields: { key: val } }
    // ============================================================
    if (data.actionItem && data.actionFields) {
      return handleActionEdit(ss, data.actionItem, data.actionFields);
    }

    // ============================================================
    // ACTION ITEM: Create new action item
    // Payload: { newAction: { priority, action, facility, ... } }
    // ============================================================
    if (data.newAction) {
      return handleNewAction(ss, data.newAction);
    }

    // ============================================================
    // ACTION ITEM: Toggle done status
    // Payload: { actionDone: "text", done: true/false }
    // ============================================================
    if (data.actionDone !== undefined) {
      return handleActionDone(ss, data.actionDone, data.done);
    }

    // ============================================================
    // PIPELINE: Create new deal
    // Payload: { newDeal: { name, status, priority, type, states, ... } }
    // ============================================================
    if (data.newDeal) {
      return handleNewDeal(ss, data.newDeal);
    }

    // ============================================================
    // PIPELINE: Deal updates (existing functionality)
    // ============================================================
    var dealName = data.dealName;
    var sheet = ss.getSheetByName('Pipeline Dashboard');

    if (!sheet) {
      return jsonResponse({ success: false, error: 'Sheet "Pipeline Dashboard" not found' });
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
      return jsonResponse({ success: false, error: 'Could not find "Facility Name" header' });
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
      return jsonResponse({ success: false, error: 'Deal "' + dealName + '" not found in sheet' });
    }

    // Legacy path: status-only update (from drag-and-drop)
    if (data.newStatus && !data.fields) {
      var statusCol = colMap['Status'];
      if (statusCol === undefined) {
        return jsonResponse({ success: false, error: 'Could not find "Status" column' });
      }
      sheet.getRange(dealRow + 1, statusCol + 1).setValue(data.newStatus);
      return jsonResponse({ success: true, deal: dealName, newStatus: data.newStatus });
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
      return jsonResponse({ success: true, deal: dealName, updated: updatedFields });
    }

    return jsonResponse({ success: false, error: 'Missing newStatus or fields' });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ============================================================
// Helper: JSON response
// ============================================================
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// Helper: Get Action Items sheet info (header row, column map)
// ============================================================
function getActionSheet(ss) {
  var sheet = ss.getSheetByName(ACTION_ITEMS_TAB);
  if (!sheet) return null;

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  // Find header row containing "Action Item" and "Priority"
  var headerRow = -1;
  for (var r = 0; r < Math.min(values.length, 10); r++) {
    var row = values[r].map(function(c) { return String(c).trim(); });
    if (row.indexOf('Action Item') >= 0 && row.indexOf('Priority') >= 0) {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) return null;

  var colMap = {};
  for (var c = 0; c < values[headerRow].length; c++) {
    colMap[String(values[headerRow][c]).trim()] = c;
  }

  return { sheet: sheet, values: values, headerRow: headerRow, colMap: colMap };
}

// ============================================================
// Handle action item edit: find row by "Action Item" text, update fields
// ============================================================
function handleActionEdit(ss, actionItemText, fields) {
  var info = getActionSheet(ss);
  if (!info) return jsonResponse({ success: false, error: 'Action Items tab not found' });

  var actionCol = info.colMap['Action Item'];
  if (actionCol === undefined) return jsonResponse({ success: false, error: 'Action Item column not found' });

  // Find the row matching the action item text
  var targetRow = -1;
  for (var r = info.headerRow + 1; r < info.values.length; r++) {
    if (String(info.values[r][actionCol]).trim() === actionItemText) {
      targetRow = r;
      break;
    }
  }

  if (targetRow < 0) {
    return jsonResponse({ success: false, error: 'Action item not found: "' + actionItemText + '"' });
  }

  var updatedFields = [];
  for (var fieldKey in fields) {
    var header = ACTION_FIELD_TO_HEADER[fieldKey];
    if (!header || info.colMap[header] === undefined) continue;
    info.sheet.getRange(targetRow + 1, info.colMap[header] + 1).setValue(fields[fieldKey]);
    updatedFields.push(fieldKey);
  }

  return jsonResponse({ success: true, action: 'edit', updated: updatedFields });
}

// ============================================================
// Handle new action item: append row to Action Items tab
// ============================================================
function handleNewAction(ss, actionData) {
  var info = getActionSheet(ss);
  if (!info) return jsonResponse({ success: false, error: 'Action Items tab not found' });

  // Build a new row array matching the header columns
  var newRow = [];
  for (var c = 0; c < info.values[info.headerRow].length; c++) {
    newRow.push(''); // initialize empty
  }

  // Fill in values from actionData using the mapping
  for (var fieldKey in actionData) {
    var header = ACTION_FIELD_TO_HEADER[fieldKey];
    if (header && info.colMap[header] !== undefined) {
      newRow[info.colMap[header]] = actionData[fieldKey];
    }
  }

  // Append to the sheet
  info.sheet.appendRow(newRow);

  return jsonResponse({ success: true, action: 'create' });
}

// ============================================================
// Handle action done toggle: update Status column to "Done" or previous value
// ============================================================
function handleActionDone(ss, actionItemText, isDone) {
  var info = getActionSheet(ss);
  if (!info) return jsonResponse({ success: false, error: 'Action Items tab not found' });

  var actionCol = info.colMap['Action Item'];
  var statusCol = info.colMap['Status'];
  if (actionCol === undefined || statusCol === undefined) {
    return jsonResponse({ success: false, error: 'Required columns not found' });
  }

  // Find the row
  var targetRow = -1;
  for (var r = info.headerRow + 1; r < info.values.length; r++) {
    if (String(info.values[r][actionCol]).trim() === actionItemText) {
      targetRow = r;
      break;
    }
  }

  if (targetRow < 0) {
    return jsonResponse({ success: false, error: 'Action item not found' });
  }

  // Set status to "Done" or restore to "Pending" when un-done
  var newStatus = isDone ? 'Done' : 'Pending';
  info.sheet.getRange(targetRow + 1, statusCol + 1).setValue(newStatus);

  return jsonResponse({ success: true, action: 'done', done: isDone });
}

// ============================================================
// Handle new deal: append row to Pipeline Dashboard tab
// ============================================================
function handleNewDeal(ss, dealData) {
  var sheet = ss.getSheetByName('Pipeline Dashboard');
  if (!sheet) return jsonResponse({ success: false, error: 'Pipeline Dashboard tab not found' });

  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();

  // Find header row
  var headerRow = -1;
  for (var r = 0; r < Math.min(values.length, 10); r++) {
    for (var c = 0; c < values[r].length; c++) {
      if (String(values[r][c]).trim() === 'Facility Name') {
        headerRow = r;
        break;
      }
    }
    if (headerRow >= 0) break;
  }
  if (headerRow < 0) return jsonResponse({ success: false, error: 'Could not find header row' });

  // Build column map
  var colMap = {};
  for (var c = 0; c < values[headerRow].length; c++) {
    colMap[String(values[headerRow][c]).trim()] = c;
  }

  // Build new row
  var newRow = [];
  for (var c = 0; c < values[headerRow].length; c++) {
    newRow.push('');
  }

  // Map deal data to columns
  var mapping = {
    'Facility Name': dealData.name || '',
    'Status': dealData.status || 'New Lead',
    'Priority': dealData.priority || '2 - Medium',
    'Type': dealData.type || '',
    'State(s)': dealData.states || '',
    'Key Contact': dealData.keyContact || '',
    'Notes': dealData.notes || '',
    'EBITDA / Financials': dealData.ebitda || '',
    'Asking Price': dealData.askingPrice || '',
    'NDA Status': dealData.ndaStatus || 'N/A',
    'Data Room': dealData.dataRoom || '',
    'Site Visit': dealData.siteVisit || '',
    'Next Action': dealData.nextAction || '',
    'Action Owner': dealData.actionOwner || '',
    'Deadline': dealData.deadline || '',
    'Last Update': dealData.lastUpdate || '',
    'Days Since Update': dealData.daysSinceUpdate || '0',
    '#': dealData.dealNumber || '',
    'Airtable ID': dealData.airtableId || ''
  };

  for (var header in mapping) {
    if (colMap[header] !== undefined) {
      newRow[colMap[header]] = mapping[header];
    }
  }

  // Append the row
  sheet.appendRow(newRow);

  return jsonResponse({ success: true, action: 'newDeal', deal: dealData.name });
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
