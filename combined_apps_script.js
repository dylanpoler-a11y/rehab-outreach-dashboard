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

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'ok',
    message: 'Rehab Outreach Pipeline API is running',
    routes: ['updateStatus (POST)', 'updateField (POST)', 'updateAction (POST)', 'newDeal (POST)']
  })).setMimeType(ContentService.MimeType.JSON);
}

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
    // SLIDES: Create Google Slides executive summary
    // Payload: { createSlides: { stateName, summary, pipeline, weeklyData, ... } }
    // ============================================================
    if (data.createSlides) {
      return handleCreateSlides(data.createSlides);
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

// ============================================================
// SLIDES: Create Google Slides executive summary for a state
// ============================================================

function handleCreateSlides(data) {
  try {
    var pres = SlidesApp.create('State Outreach Report \u2014 ' + data.stateName + ' (' + data.generatedDate + ')');

    // Slide 1: Title
    buildTitleSlide(pres, data);

    // Slide 2: Executive overview
    buildOverviewSlide(pres, data);

    // Slide 3: Weekly outreach chart (with scheduled call grouping)
    if (data.weeklyData && data.weeklyData.length > 0) {
      buildWeeklyChartSlide(pres, data);
    }

    // Slide 4: Pipeline details (only if deals exist)
    if (data.pipeline && data.pipeline.totalDeals > 0) {
      buildPipelineSlide(pres, data);
    }

    pres.saveAndClose();

    return jsonResponse({
      success: true,
      url: pres.getUrl(),
      slideCount: pres.getSlides().length
    });
  } catch (err) {
    return jsonResponse({ success: false, error: 'Slides creation failed: ' + err.toString() });
  }
}

// --- Slide helpers ---

function addSlideTitle(slide, text) {
  var box = slide.insertTextBox(text, 40, 20, 640, 36);
  box.getText().getTextStyle()
    .setFontFamily('Inter')
    .setFontSize(22)
    .setBold(true)
    .setForegroundColor('#f1f5f9');
}

function addSlideSubheading(slide, text, x, y) {
  var box = slide.insertTextBox(text, x, y, 300, 20);
  box.getText().getTextStyle()
    .setFontFamily('Inter')
    .setFontSize(12)
    .setBold(true)
    .setForegroundColor('#cbd5e1');
}

// --- Slide 1: Title ---

function buildTitleSlide(pres, data) {
  var slide = pres.getSlides()[0];
  slide.getBackground().setSolidFill('#0f172a');

  // Clear default placeholders
  var elements = slide.getPageElements();
  for (var i = elements.length - 1; i >= 0; i--) {
    elements[i].remove();
  }

  // Title
  var title = slide.insertTextBox('State Outreach Report', 50, 100, 620, 60);
  title.getText().getTextStyle()
    .setFontFamily('Inter')
    .setFontSize(36)
    .setBold(true)
    .setForegroundColor('#f1f5f9');

  // State name
  var stateBox = slide.insertTextBox(data.stateName, 50, 170, 620, 70);
  stateBox.getText().getTextStyle()
    .setFontFamily('Inter')
    .setFontSize(48)
    .setBold(true)
    .setForegroundColor('#3b82f6');

  // Date + summary
  var s = data.summary || {};
  var subtitle = data.generatedDate + '  \u2022  ' + (s.total || 0) + ' Companies Tracked  \u2022  ' + (s.totalMsgs || 0) + ' Messages Sent';
  var dateBox = slide.insertTextBox(subtitle, 50, 260, 620, 30);
  dateBox.getText().getTextStyle()
    .setFontFamily('Inter')
    .setFontSize(14)
    .setForegroundColor('#94a3b8');

  // Accent line
  var line = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 50, 310, 200, 4);
  line.getFill().setSolidFill('#3b82f6');
  line.getBorder().setTransparent();

  // Branding
  var brand = slide.insertTextBox('Behavioral Health Outreach Hub', 50, 350, 400, 20);
  brand.getText().getTextStyle()
    .setFontFamily('Inter')
    .setFontSize(10)
    .setForegroundColor('#475569');
}

// --- Slide 2: Executive Overview ---

function buildOverviewSlide(pres, data) {
  var slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill('#0f172a');

  addSlideTitle(slide, 'Executive Overview');

  var s = data.summary || {};

  // --- TOP ROW: 4 metric cards ---
  var metrics = [
    { label: 'Total Companies', value: String(s.total || 0), color: '#3b82f6' },
    { label: "Mom 'n Pops", value: String(s.momNPop || 0), color: '#10b981' },
    { label: 'Contacted', value: String(s.contacted || 0), color: '#f59e0b' },
    { label: 'Messages Sent', value: String(s.totalMsgs || 0), color: '#06b6d4' }
  ];

  var cardW = 145, cardGap = 15, startX = 40, cardY = 66;
  for (var i = 0; i < metrics.length; i++) {
    var m = metrics[i];
    var x = startX + i * (cardW + cardGap);
    var card = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, x, cardY, cardW, 72);
    card.getFill().setSolidFill('#1e293b');
    card.getBorder().getLineFill().setSolidFill('#334155');
    card.getBorder().setWeight(1);

    var valBox = slide.insertTextBox(m.value, x + 10, cardY + 8, cardW - 20, 36);
    valBox.getText().getTextStyle().setFontFamily('Inter').setFontSize(26).setBold(true).setForegroundColor(m.color);

    var lblBox = slide.insertTextBox(m.label, x + 10, cardY + 44, cardW - 20, 18);
    lblBox.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#94a3b8');
  }

  // --- MIDDLE: Outreach funnel (Mom 'n Pop only) ---
  var funnelY = 155;
  addSlideSubheading(slide, "Outreach Funnel — Mom 'n Pop (" + (s.momNPop || 0) + ' companies)', 40, funnelY);

  var mnpContacted = s.mnpContacted || 0;
  var mnpResponded = s.mnpResponded || 0;
  var mnpNotResponded = s.mnpNotResponded || 0;
  var mnpMeetings = s.mnpAssistedMeeting || 0;
  var mnpPipeline = s.mnpInPipeline || 0;
  var mnpBase = s.momNPop || 1;

  var funnelSteps = [
    { label: 'Not Responded', value: String(mnpNotResponded), pct: ((mnpNotResponded / mnpBase) * 100).toFixed(1) + '%', color: '#64748b' },
    { label: 'Responded', value: String(mnpResponded), pct: ((mnpResponded / mnpBase) * 100).toFixed(1) + '%', color: '#3b82f6' },
    { label: 'Meetings', value: String(mnpMeetings), pct: ((mnpMeetings / mnpBase) * 100).toFixed(1) + '%', color: '#10b981' },
    { label: 'In Pipeline', value: String(mnpPipeline), pct: ((mnpPipeline / mnpBase) * 100).toFixed(1) + '%', color: '#ec4899' }
  ];

  var funnelItemY = funnelY + 24;
  var fw = 140;
  for (var i = 0; i < funnelSteps.length; i++) {
    var step = funnelSteps[i];
    var fx = 40 + i * (fw + 18);
    var rect = slide.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, fx, funnelItemY, fw, 60);
    rect.getFill().setSolidFill('#1e293b');
    rect.getBorder().getLineFill().setSolidFill(step.color);
    rect.getBorder().setWeight(2);

    var vb = slide.insertTextBox(step.value, fx + 10, funnelItemY + 2, fw - 20, 22);
    vb.getText().getTextStyle().setFontFamily('Inter').setFontSize(18).setBold(true).setForegroundColor(step.color);

    var pctBox = slide.insertTextBox(step.pct, fx + 10, funnelItemY + 24, fw - 20, 14);
    pctBox.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor(step.color);

    var lb = slide.insertTextBox(step.label, fx + 10, funnelItemY + 40, fw - 20, 14);
    lb.getText().getTextStyle().setFontFamily('Inter').setFontSize(8).setForegroundColor('#94a3b8');

    // Arrow between steps
    if (i < funnelSteps.length - 1) {
      var arrowBox = slide.insertTextBox('\u2192', fx + fw + 2, funnelItemY + 16, 14, 20);
      arrowBox.getText().getTextStyle().setFontFamily('Inter').setFontSize(16).setForegroundColor('#475569');
    }
  }

  // --- NDA / LOI / EBITDA summary ---
  var p = data.pipeline || {};
  var ndaY = 258;
  addSlideSubheading(slide, 'Pipeline & Deal Progress', 40, ndaY);

  var ndaText = 'Pipeline Deals: ' + (p.totalDeals || 0) +
    '  |  NDAs Signed: ' + (p.ndaSigned || 0) +
    '  |  NDAs Sent: ' + (p.ndaSent || 0) +
    '  |  LOIs: ' + (p.loiCount || 0);
  var ndaBox = slide.insertTextBox(ndaText, 40, ndaY + 20, 640, 18);
  ndaBox.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setForegroundColor('#cbd5e1');

  // LOI details
  if (p.loiDeals && p.loiDeals.length > 0) {
    var loiLines = p.loiDeals.map(function(d) { return d.name + (d.askingPrice ? ' (' + d.askingPrice + ')' : ''); }).join('  |  ');
    var loiBox = slide.insertTextBox('LOI Details: ' + loiLines, 40, ndaY + 40, 640, 16);
    loiBox.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#94a3b8');
  }

  // EBITDA details
  if (p.ebitdaDeals && p.ebitdaDeals.length > 0) {
    var ey = (p.loiDeals && p.loiDeals.length > 0) ? ndaY + 58 : ndaY + 40;
    var ebitdaLines = p.ebitdaDeals.map(function(d) { return d.name + ': ' + d.ebitda; }).join('  |  ');
    var ebitdaBox = slide.insertTextBox('EBITDA: ' + ebitdaLines, 40, ey, 640, 16);
    ebitdaBox.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#94a3b8');
  }

  // --- Ownership breakdown bar ---
  var ownerY = 330;
  addSlideSubheading(slide, 'Ownership Breakdown', 40, ownerY);

  var ownerColors = {
    "Mom 'n Pop": '#10b981', 'Private Equity': '#8b5cf6',
    'Publicly Traded': '#3b82f6', 'Non-Profit': '#f59e0b',
    'Bradford Facility': '#06b6d4', 'Bradford OP Office': '#14b8a6',
    'CLOSED': '#ef4444', 'GOV': '#64748b'
  };
  var barX = 40, barTotalW = 640, barH = 22, barStartY = ownerY + 22;
  var total = s.total || 1;
  var oc = data.ownershipCounts || {};
  var keys = Object.keys(oc);

  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var count = oc[key];
    var w = Math.max(Math.round((count / total) * barTotalW), 4);
    var color = ownerColors[key] || '#64748b';

    var seg = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, barX, barStartY, w, barH);
    seg.getFill().setSolidFill(color);
    seg.getBorder().setTransparent();

    // Label below for wider segments
    if (w > 40) {
      var segLabel = slide.insertTextBox(key + ' (' + count + ')', barX, barStartY + barH + 2, w, 14);
      segLabel.getText().getTextStyle().setFontFamily('Inter').setFontSize(7).setForegroundColor('#94a3b8');
      segLabel.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    }

    barX += w;
  }
}

// --- Slide 4: Pipeline Detail Table ---

function buildPipelineSlide(pres, data) {
  var slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill('#0f172a');

  addSlideTitle(slide, 'Pipeline Details \u2014 ' + data.stateName);

  var details = (data.pipeline && data.pipeline.details) || [];
  var cols = ['Facility', 'Status', 'NDA', 'EBITDA', 'Asking Price', 'Priority'];
  var colWidths = [160, 100, 70, 100, 90, 80];
  var tableX = 30, tableY = 66;
  var headerH = 30, rowH = 26;

  // Header row
  var hx = tableX;
  for (var i = 0; i < cols.length; i++) {
    var hdr = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, hx, tableY, colWidths[i], headerH);
    hdr.getFill().setSolidFill('#1e293b');
    hdr.getBorder().getLineFill().setSolidFill('#334155');
    hdr.getBorder().setWeight(1);

    var ht = slide.insertTextBox(cols[i], hx + 6, tableY + 6, colWidths[i] - 12, headerH - 12);
    ht.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setBold(true).setForegroundColor('#94a3b8');

    hx += colWidths[i];
  }

  // Data rows
  var maxRows = Math.min(details.length, 12);
  for (var r = 0; r < maxRows; r++) {
    var deal = details[r];
    var rowValues = [
      deal.name || '-',
      deal.status || '-',
      deal.ndaStatus || '-',
      deal.ebitda || 'TBD',
      deal.askingPrice || 'TBD',
      deal.priority || '-'
    ];
    var cx = tableX;
    var cy = tableY + headerH + (r * rowH);
    var bgColor = r % 2 === 0 ? '#0f172a' : '#1a2332';

    for (var c = 0; c < rowValues.length; c++) {
      var cell = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, cx, cy, colWidths[c], rowH);
      cell.getFill().setSolidFill(bgColor);
      cell.getBorder().getLineFill().setSolidFill('#1e293b');
      cell.getBorder().setWeight(0.5);

      var ct = slide.insertTextBox(String(rowValues[c]), cx + 6, cy + 4, colWidths[c] - 12, rowH - 8);
      ct.getText().getTextStyle().setFontFamily('Inter').setFontSize(8).setForegroundColor('#e2e8f0');

      cx += colWidths[c];
    }
  }

  // Overflow note
  if (details.length > maxRows) {
    var noteY = tableY + headerH + (maxRows * rowH) + 8;
    var note = slide.insertTextBox('+ ' + (details.length - maxRows) + ' additional deals not shown', tableX, noteY, 400, 16);
    note.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setItalic(true).setForegroundColor('#64748b');
  }
}

// --- Slide 3: Weekly Outreach Bar Chart (with Scheduled Call grouping) ---

function buildWeeklyChartSlide(pres, data) {
  var slide = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  slide.getBackground().setSolidFill('#0f172a');

  addSlideTitle(slide, 'Weekly Outreach Activity \u2014 ' + data.stateName);

  var weeks = data.weeklyData || [];
  if (weeks.length === 0) return;

  var chartX = 60, chartY = 80, chartW = 600, chartH = 240;
  var barGap = 6;
  var numBars = weeks.length;
  var barW = Math.floor((chartW - (numBars - 1) * barGap) / numBars);
  var maxCount = 1;
  for (var i = 0; i < weeks.length; i++) {
    if (weeks[i].count > maxCount) maxCount = weeks[i].count;
  }

  // Baseline
  var baseY = chartY + chartH;
  var baseline = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, chartX - 2, baseY, chartW + 4, 2);
  baseline.getFill().setSolidFill('#334155');
  baseline.getBorder().setTransparent();

  // Stacked bars: bottom = Scheduled Call (green), top = No Scheduled Call (blue)
  for (var i = 0; i < weeks.length; i++) {
    var w = weeks[i];
    var totalH = Math.max(Math.round((w.count / maxCount) * (chartH - 40)), 4);
    var bx = chartX + i * (barW + barGap);

    var scheduledCount = w.scheduled || 0;
    var notScheduledCount = w.notScheduled || 0;

    // Calculate segment heights proportionally
    var scheduledH = w.count > 0 ? Math.round((scheduledCount / w.count) * totalH) : 0;
    var notScheduledH = totalH - scheduledH;

    // Draw not-scheduled (blue) segment on top
    if (notScheduledH > 0) {
      var topY = baseY - totalH;
      var topBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, bx, topY, barW, notScheduledH);
      topBar.getFill().setSolidFill('#3b82f6');
      topBar.getBorder().setTransparent();
    }

    // Draw scheduled (green) segment on bottom
    if (scheduledH > 0) {
      var botY = baseY - scheduledH;
      var botBar = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, bx, botY, barW, scheduledH);
      botBar.getFill().setSolidFill('#10b981');
      botBar.getBorder().setTransparent();
    }

    // Count label above bar showing breakdown
    var countLabel = String(w.count);
    if (scheduledCount > 0) {
      countLabel = scheduledCount + '/' + w.count;
    }
    var countBox = slide.insertTextBox(countLabel, bx - 2, baseY - totalH - 16, barW + 4, 14);
    countBox.getText().getTextStyle().setFontFamily('Inter').setFontSize(7).setBold(true).setForegroundColor('#93c5fd');
    countBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    // Week label below
    var labelBox = slide.insertTextBox(w.label || '', bx - 4, baseY + 4, barW + 8, 14);
    labelBox.getText().getTextStyle().setFontFamily('Inter').setFontSize(7).setForegroundColor('#64748b');
    labelBox.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }

  // Y-axis label
  var yLabel = slide.insertTextBox('Messages', 8, chartY + chartH / 2 - 10, 46, 18);
  yLabel.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#64748b');

  // Legend
  var legendY = baseY + 22;
  var legendGreen = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 60, legendY, 12, 12);
  legendGreen.getFill().setSolidFill('#10b981');
  legendGreen.getBorder().setTransparent();
  var legendGreenLabel = slide.insertTextBox('Scheduled Call', 76, legendY - 1, 100, 14);
  legendGreenLabel.getText().getTextStyle().setFontFamily('Inter').setFontSize(8).setForegroundColor('#94a3b8');

  var legendBlue = slide.insertShape(SlidesApp.ShapeType.RECTANGLE, 180, legendY, 12, 12);
  legendBlue.getFill().setSolidFill('#3b82f6');
  legendBlue.getBorder().setTransparent();
  var legendBlueLabel = slide.insertTextBox('No Scheduled Call', 196, legendY - 1, 120, 14);
  legendBlueLabel.getText().getTextStyle().setFontFamily('Inter').setFontSize(8).setForegroundColor('#94a3b8');

  // Medium breakdown note at bottom
  var mc = data.mediumCounts || {};
  var medKeys = Object.keys(mc);
  if (medKeys.length > 0) {
    var medText = 'By Channel: ' + medKeys.map(function(k) { return k + ' (' + mc[k] + ')'; }).join('  \u2022  ');
    var medBox = slide.insertTextBox(medText, 40, legendY + 18, 640, 16);
    medBox.getText().getTextStyle().setFontFamily('Inter').setFontSize(8).setForegroundColor('#64748b');
  }
}

function createPitchDeck() {
  var pres = SlidesApp.create('The Poler Team \u2014 M&A Lead Generation');
  var firstSlide = pres.getSlides()[0];
  firstSlide.remove();

  // Color palette
  var DARK = '#0f172a';
  var ACCENT = '#2563eb';
  var GREEN = '#10b981';
  var WHITE = '#ffffff';
  var GRAY = '#94a3b8';
  var LIGHT_BG = '#1e293b';
  var GOLD = '#f59e0b';

  // ==================== SLIDE 1: TITLE ====================
  var s1 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s1.getBackground().setSolidFill(DARK);

  var title1 = s1.insertTextBox('The Poler Team', 40, 140, 640, 60);
  title1.getText().getTextStyle().setFontFamily('Inter').setFontSize(42).setBold(true).setForegroundColor(WHITE);

  var sub1 = s1.insertTextBox('M&A Lead Generation', 40, 210, 640, 40);
  sub1.getText().getTextStyle().setFontFamily('Inter').setFontSize(24).setForegroundColor(ACCENT);

  var line1 = s1.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 270, 200, 4);
  line1.getFill().setSolidFill(ACCENT);
  line1.getBorder().setTransparent();

  var desc1 = s1.insertTextBox('Connecting buyers with off-market small business owners\nwho are not going through a formal sale process.', 40, 300, 600, 60);
  desc1.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setForegroundColor(GRAY);

  var prep = s1.insertTextBox('Prepared for Olympus Cosmetic Group', 40, 390, 400, 24);
  prep.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setForegroundColor(GOLD);

  var footer1 = s1.insertTextBox('Aventura, FL  \u2022  thepolerteam.com  \u2022  Confidential', 40, 490, 400, 20);
  footer1.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor(GRAY);

  // ==================== SLIDE 2: WHO WE ARE ====================
  var s2 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s2.getBackground().setSolidFill(DARK);

  var h2 = s2.insertTextBox('Who We Are', 40, 30, 640, 40);
  h2.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line2 = s2.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line2.getFill().setSolidFill(ACCENT);
  line2.getBorder().setTransparent();

  var body2 = s2.insertTextBox(
    'The Poler Team is a family-operated firm based in Aventura, Florida with over 20 years of real estate experience.\n\n' +
    'Rosa Poler founded the team in residential real estate, expanding into commercial transactions. Today, Dylan Poler leads our M&A Lead Generation division \u2014 focused exclusively on connecting acquisition-minded buyers with privately-held small business owners who are not listed on the market.\n\n' +
    'We do not represent multiple competitors in the same industry. Each buyer engagement is exclusive within their sector, ensuring confidentiality and alignment of interests.\n\n' +
    'Our approach is direct, personal outreach to owner-operators \u2014 not mass marketing. We identify, contact, and qualify small business owners, then facilitate introductory calls between owners and our buyer clients.',
    40, 95, 640, 240
  );
  body2.getText().getTextStyle().setFontFamily('Inter').setFontSize(12).setForegroundColor('#cbd5e1');
  body2.getText().getParagraphStyle().setLineSpacing(130);

  // Value props
  var vp1 = s2.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 360, 200, 80);
  vp1.getFill().setSolidFill(LIGHT_BG);
  vp1.getBorder().getLineFill().setSolidFill('#334155');
  var vt1 = s2.insertTextBox('Off-Market\nDeal Sourcing', 50, 375, 180, 50);
  vt1.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setBold(true).setForegroundColor(WHITE);
  vt1.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  var vp2 = s2.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 260, 360, 200, 80);
  vp2.getFill().setSolidFill(LIGHT_BG);
  vp2.getBorder().getLineFill().setSolidFill('#334155');
  var vt2 = s2.insertTextBox('Industry\nExclusivity', 270, 375, 180, 50);
  vt2.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setBold(true).setForegroundColor(WHITE);
  vt2.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  var vp3 = s2.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 480, 360, 200, 80);
  vp3.getFill().setSolidFill(LIGHT_BG);
  vp3.getBorder().getLineFill().setSolidFill('#334155');
  var vt3 = s2.insertTextBox('No Bidding\nProcess', 490, 375, 180, 50);
  vt3.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setBold(true).setForegroundColor(WHITE);
  vt3.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  var footer2 = s2.insertTextBox('The Poler Team  \u2022  Confidential', 40, 500, 300, 16);
  footer2.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 3: THE PROBLEM WE SOLVE ====================
  var s3 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s3.getBackground().setSolidFill(DARK);

  var h3 = s3.insertTextBox('The Problem We Solve', 40, 30, 640, 40);
  h3.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line3 = s3.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line3.getFill().setSolidFill(ACCENT);
  line3.getBorder().setTransparent();

  // Left column - The Challenge
  var ch3 = s3.insertTextBox('The Challenge', 40, 95, 310, 24);
  ch3.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setBold(true).setForegroundColor(GOLD);

  var chBody = s3.insertTextBox(
    '\u2022 Most small business owners are not actively listing their businesses for sale\n\n' +
    '\u2022 Brokers and investment bankers create competitive bidding, driving up prices\n\n' +
    '\u2022 Off-market owners are difficult to identify and even harder to engage\n\n' +
    '\u2022 Cold outreach at scale requires systems, persistence, and a personal touch',
    40, 125, 310, 200
  );
  chBody.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setForegroundColor('#cbd5e1');
  chBody.getText().getParagraphStyle().setLineSpacing(120);

  // Right column - Our Solution
  var sol3 = s3.insertTextBox('Our Solution', 380, 95, 310, 24);
  sol3.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setBold(true).setForegroundColor(GREEN);

  var solBody = s3.insertTextBox(
    '\u2022 We directly contact owner-operators across the country via personalized outreach\n\n' +
    '\u2022 We achieve a 33.7% response rate from small business owners \u2014 far above industry norms\n\n' +
    '\u2022 75% of respondents agree to an introductory call\n\n' +
    '\u2022 No broker, no bidding war \u2014 just a direct conversation between buyer and seller',
    380, 125, 310, 200
  );
  solBody.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setForegroundColor('#cbd5e1');
  solBody.getText().getParagraphStyle().setLineSpacing(120);

  // Bottom highlight box
  var hlBox = s3.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, 350, 640, 80);
  hlBox.getFill().setSolidFill('#1a2744');
  hlBox.getBorder().getLineFill().setSolidFill(ACCENT);
  var hlTxt = s3.insertTextBox('In our most recent engagement, we sourced and closed a $20M acquisition\nwhere the seller did not use an investment banker or broker \u2014\nno bidding process, direct deal.', 60, 360, 600, 60);
  hlTxt.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setBold(true).setForegroundColor(WHITE);
  hlTxt.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  var footer3 = s3.insertTextBox('The Poler Team  \u2022  Confidential', 40, 500, 300, 16);
  footer3.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 4: PROVEN RESULTS ====================
  var s4 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s4.getBackground().setSolidFill(DARK);

  var h4 = s4.insertTextBox('Proven Results at Scale', 40, 30, 640, 40);
  h4.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line4 = s4.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line4.getFill().setSolidFill(ACCENT);
  line4.getBorder().setTransparent();

  var sub4 = s4.insertTextBox('Results from a single client engagement in the behavioral health sector', 40, 85, 600, 20);
  sub4.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setForegroundColor(GRAY);

  // Metric cards - Row 1
  var cards = [
    { val: '500+', label: 'Small Business Owners\nContacted', color: ACCENT },
    { val: '33.7%', label: 'Response Rate\nfrom Owners', color: GREEN },
    { val: '75.1%', label: 'Of Respondents\nScheduled Intro Calls', color: '#8b5cf6' },
    { val: '86.5%', label: 'Of Intro Calls Led to\nAssisted Meetings', color: GOLD }
  ];

  for (var i = 0; i < cards.length; i++) {
    var cx = 40 + i * 168;
    var card = s4.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, cx, 115, 155, 95);
    card.getFill().setSolidFill(LIGHT_BG);
    card.getBorder().getLineFill().setSolidFill(cards[i].color);

    var cv = s4.insertTextBox(cards[i].val, cx + 10, 122, 135, 36);
    cv.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(cards[i].color);

    var cl = s4.insertTextBox(cards[i].label, cx + 10, 162, 135, 40);
    cl.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor(GRAY);
  }

  // Funnel visualization
  var funnelH = s4.insertTextBox('Outreach Funnel \u2014 Owner-Operators Only', 40, 225, 400, 20);
  funnelH.getText().getTextStyle().setFontFamily('Inter').setFontSize(12).setBold(true).setForegroundColor(WHITE);

  var funnelSteps = [
    { val: 526, label: 'Contacted', w: 580, color: '#334155' },
    { val: 177, label: 'Responded (33.7%)', w: 430, color: ACCENT },
    { val: 133, label: 'Intro Call Scheduled (75.1%)', w: 320, color: '#7c3aed' },
    { val: 115, label: 'Assisted Meeting (86.5%)', w: 260, color: GREEN },
    { val: 16, label: 'In Active Pipeline', w: 120, color: GOLD }
  ];

  for (var j = 0; j < funnelSteps.length; j++) {
    var fy = 252 + j * 38;
    var bar = s4.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 80, fy, funnelSteps[j].w, 28);
    bar.getFill().setSolidFill(funnelSteps[j].color);
    bar.getBorder().setTransparent();

    var fLabel = s4.insertTextBox(funnelSteps[j].val + '  ' + funnelSteps[j].label, 90, fy + 4, funnelSteps[j].w - 20, 20);
    fLabel.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setBold(true).setForegroundColor(WHITE);
  }

  // Additional stats
  var addStats = s4.insertTextBox(
    'Coverage: 35+ states  \u2022  Active pipeline deals across multiple states  \u2022  $20M closed deal (no broker, no bidding)',
    40, 460, 640, 20
  );
  addStats.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor(GRAY);

  var footer4 = s4.insertTextBox('The Poler Team  \u2022  Confidential  \u2022  Sanitized data from a single client engagement', 40, 500, 500, 16);
  footer4.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 5: HOW WE WORK ====================
  var s5 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s5.getBackground().setSolidFill(DARK);

  var h5 = s5.insertTextBox('How We Work', 40, 30, 640, 40);
  h5.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line5 = s5.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line5.getFill().setSolidFill(ACCENT);
  line5.getBorder().setTransparent();

  var steps = [
    { num: '01', title: 'Research & Targeting', desc: 'We identify owner-operated businesses that fit your acquisition criteria \u2014 by geography, size, specialty, and ownership type.' },
    { num: '02', title: 'Personalized Outreach', desc: 'Our team contacts owners directly via LinkedIn, phone, email, text, and voicemail. Every message is personalized. No mass blasts.' },
    { num: '03', title: 'Qualify & Schedule', desc: 'When an owner expresses interest, we qualify the opportunity and schedule an introductory call between you and the owner. No commitment required from either party.' },
    { num: '04', title: 'Facilitate & Support', desc: 'We assist through the meeting process and remain a resource as the relationship develops \u2014 from first call through LOI and beyond.' }
  ];

  for (var k = 0; k < steps.length; k++) {
    var sy = 95 + k * 100;
    var numBox = s5.insertShape(SlidesApp.ShapeType.ELLIPSE, 40, sy + 5, 40, 40);
    numBox.getFill().setSolidFill(ACCENT);
    numBox.getBorder().setTransparent();
    var numTxt = s5.insertTextBox(steps[k].num, 40, sy + 12, 40, 24);
    numTxt.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setBold(true).setForegroundColor(WHITE);
    numTxt.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

    var stTitle = s5.insertTextBox(steps[k].title, 95, sy, 585, 24);
    stTitle.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setBold(true).setForegroundColor(WHITE);

    var stDesc = s5.insertTextBox(steps[k].desc, 95, sy + 26, 585, 50);
    stDesc.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setForegroundColor('#cbd5e1');
    stDesc.getText().getParagraphStyle().setLineSpacing(120);
  }

  var footer5 = s5.insertTextBox('The Poler Team  \u2022  Confidential', 40, 500, 300, 16);
  footer5.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 6: WHY OFF-MARKET ====================
  var s6 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s6.getBackground().setSolidFill(DARK);

  var h6 = s6.insertTextBox('Why Off-Market Deals', 40, 30, 640, 40);
  h6.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line6 = s6.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line6.getFill().setSolidFill(ACCENT);
  line6.getBorder().setTransparent();

  // Comparison table header
  var thBg = s6.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 100, 640, 30);
  thBg.getFill().setSolidFill('#1e293b');
  thBg.getBorder().setTransparent();

  var th1 = s6.insertTextBox('', 40, 103, 200, 24);
  var th2 = s6.insertTextBox('Broker / Banker Process', 240, 103, 210, 24);
  th2.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setBold(true).setForegroundColor('#f87171');
  th2.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  var th3 = s6.insertTextBox('The Poler Team', 470, 103, 210, 24);
  th3.getText().getTextStyle().setFontFamily('Inter').setFontSize(11).setBold(true).setForegroundColor(GREEN);
  th3.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);

  var rows = [
    ['Competition', 'Multiple bidders', 'You are the only buyer'],
    ['Pricing', 'Inflated by bidding', 'Negotiate directly with owner'],
    ['Timeline', 'Lengthy formal process', 'Move at your pace'],
    ['Relationship', 'Filtered through intermediary', 'Direct owner relationship'],
    ['Deal Flow', 'Wait for listings', 'Proactive \u2014 we find them'],
    ['Confidentiality', 'Widely marketed', 'Private, targeted outreach']
  ];

  for (var r = 0; r < rows.length; r++) {
    var ry = 135 + r * 40;
    if (r % 2 === 0) {
      var rowBg = s6.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, ry, 640, 38);
      rowBg.getFill().setSolidFill('#0f172a');
      rowBg.getBorder().setTransparent();
    }
    var rc1 = s6.insertTextBox(rows[r][0], 50, ry + 8, 180, 22);
    rc1.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setBold(true).setForegroundColor(WHITE);
    var rc2 = s6.insertTextBox(rows[r][1], 245, ry + 8, 200, 22);
    rc2.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor('#94a3b8');
    rc2.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    var rc3 = s6.insertTextBox(rows[r][2], 475, ry + 8, 200, 22);
    rc3.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor(GREEN);
    rc3.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
  }

  var footer6 = s6.insertTextBox('The Poler Team  \u2022  Confidential', 40, 500, 300, 16);
  footer6.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 7: NEXT STEPS ====================
  var s7 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s7.getBackground().setSolidFill(DARK);

  var h7 = s7.insertTextBox('Next Steps', 40, 30, 640, 40);
  h7.getText().getTextStyle().setFontFamily('Inter').setFontSize(28).setBold(true).setForegroundColor(WHITE);

  var line7 = s7.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 75, 120, 3);
  line7.getFill().setSolidFill(ACCENT);
  line7.getBorder().setTransparent();

  var nextBody = s7.insertTextBox(
    'We would welcome the opportunity to support Olympus Cosmetic Group\'s acquisition strategy.\n\n' +
    'Our proposed next steps:',
    40, 95, 640, 60
  );
  nextBody.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setForegroundColor('#cbd5e1');

  var nextSteps = [
    { icon: '\u260E', title: 'Introductory Call', desc: 'A brief call to learn about your target criteria \u2014 geography, practice type, size, and deal structure preferences.' },
    { icon: '\uD83C\uDFAF', title: 'Target List Development', desc: 'We build a curated list of owner-operated practices matching your criteria across your target states.' },
    { icon: '\uD83D\uDCE8', title: 'Outreach Campaign Launch', desc: 'Personalized, multi-channel outreach begins. You receive qualified introductions \u2014 no upfront cost until engagement.' },
    { icon: '\uD83E\uDD1D', title: 'Buyer Agreement', desc: 'Once you see the quality of our pipeline, we formalize a buyer\u2019s agreement for ongoing deal sourcing.' }
  ];

  for (var n = 0; n < nextSteps.length; n++) {
    var ny = 170 + n * 75;
    var nBox = s7.insertShape(SlidesApp.ShapeType.ROUND_RECTANGLE, 40, ny, 640, 65);
    nBox.getFill().setSolidFill(LIGHT_BG);
    nBox.getBorder().getLineFill().setSolidFill('#334155');

    var nTitle = s7.insertTextBox(nextSteps[n].title, 60, ny + 8, 580, 22);
    nTitle.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setBold(true).setForegroundColor(WHITE);

    var nDesc = s7.insertTextBox(nextSteps[n].desc, 60, ny + 30, 600, 30);
    nDesc.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor(GRAY);
  }

  var footer7 = s7.insertTextBox('The Poler Team  \u2022  Confidential', 40, 500, 300, 16);
  footer7.getText().getTextStyle().setFontFamily('Inter').setFontSize(9).setForegroundColor('#475569');

  // ==================== SLIDE 8: CONTACT ====================
  var s8 = pres.appendSlide(SlidesApp.PredefinedLayout.BLANK);
  s8.getBackground().setSolidFill(DARK);

  var cTitle = s8.insertTextBox('Let\u2019s Connect', 40, 150, 640, 50);
  cTitle.getText().getTextStyle().setFontFamily('Inter').setFontSize(36).setBold(true).setForegroundColor(WHITE);

  var cLine = s8.insertShape(SlidesApp.ShapeType.RECTANGLE, 40, 210, 200, 4);
  cLine.getFill().setSolidFill(ACCENT);
  cLine.getBorder().setTransparent();

  var cName = s8.insertTextBox('Dylan Poler', 40, 240, 400, 30);
  cName.getText().getTextStyle().setFontFamily('Inter').setFontSize(18).setBold(true).setForegroundColor(WHITE);

  var cRole = s8.insertTextBox('M&A Lead Generation', 40, 270, 400, 24);
  cRole.getText().getTextStyle().setFontFamily('Inter').setFontSize(14).setForegroundColor(ACCENT);

  var cTeam = s8.insertTextBox('The Poler Team', 40, 300, 400, 24);
  cTeam.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setForegroundColor(GRAY);

  var cLoc = s8.insertTextBox('Aventura, FL', 40, 330, 400, 24);
  cLoc.getText().getTextStyle().setFontFamily('Inter').setFontSize(13).setForegroundColor(GRAY);

  var footer8 = s8.insertTextBox('Confidential  \u2022  For Olympus Cosmetic Group review only', 40, 490, 400, 20);
  footer8.getText().getTextStyle().setFontFamily('Inter').setFontSize(10).setForegroundColor('#475569');

  return pres.getUrl();
}
