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
