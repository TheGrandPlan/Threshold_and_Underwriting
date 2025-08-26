// =============================================================
// Threshold Script (Formatted & Streamlined)
//  - Scans sheet for new rows above ROI threshold (> 0.40)
//  - Creates folder + spreadsheet copy
//  - Populates key fields
//  - Calls n8n (optional) and Central script (for post-copy Script Properties)
//  - Marks checkbox + link
//  - Auto-sort removed from this module (handled elsewhere)
// =============================================================

// -------- Configuration (constants + helpers) --------
var THRESHOLD_SHEET_NAME = 'Deal Analysis Summary - Prospective Properties';
var HEADER_ROWS = 3; // rows 1-3 are headers

// Column indices (0-based for array access). Keep consistent.
var COL = {
  address: 1,       // B
  zip: 2,           // C
  askingPrice: 3,   // D
  lotSize: 4,       // E
  status: 6,        // G
  metric: 17,       // R
  processedCheck: 20, // U (checkbox)
  link: 21          // V (link)
};

// Script Properties access (property names already configured in project settings)
function getProp_(key, fallback) {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(key);
    return v != null && v !== '' ? v : fallback;
  } catch (e) { return fallback; }
}

var PARENT_FOLDER_ID = getProp_('PARENT_FOLDER_ID', 'MISSING_PARENT_FOLDER_ID');
var TEMPLATE_FILE_ID = getProp_('TEMPLATE_FILE_ID', 'MISSING_TEMPLATE_FILE_ID');
var N8N_WEBHOOK_URL = getProp_('N8N_WEBHOOK_URL', '');
var CENTRAL_URL = getProp_('CENTRAL_URL', ''); // Web App URL that sets Script Properties inside the new copy

// -------- Entry Point --------
function checkThresholdAndProcess() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(THRESHOLD_SHEET_NAME);
  if (!sheet) {
    Logger.log("Sheet '" + THRESHOLD_SHEET_NAME + "' not found. Exiting.");
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length <= HEADER_ROWS) {
    Logger.log('No data rows to process.');
    return;
  }

  Logger.log('Scanning rows ' + (HEADER_ROWS + 1) + ' to ' + data.length + '.');
  var processedThisRun = 0;

  for (var i = HEADER_ROWS; i < data.length; i++) {
    try {
      var row = data[i];
      var rawMetric = row[COL.metric];
      var metricVal = parseMetric_(rawMetric);
      if (metricVal == null || !(metricVal > 0.40)) continue; // threshold strictly > 0.40

      var status = String(row[COL.status] || '').trim().toUpperCase();
      if (status !== 'ACTIVE') continue;

      var checkboxVal = row[COL.processedCheck];
      if (isChecked_(checkboxVal)) continue; // already processed

      var address = row[COL.address];
      if (!address) {
        Logger.log('Skipping row ' + (i + 1) + ': missing address.');
        continue;
      }

      var zip = row[COL.zip];
      var lotSize = row[COL.lotSize];
      var askingPrice = row[COL.askingPrice];

      Logger.log('Processing row ' + (i + 1) + ' address="' + address + '" metric=' + metricVal);
      createFolderCopyAndCallCentralized(address, zip, i, lotSize, askingPrice);
      processedThisRun++;
    } catch (err) {
      Logger.log('Error row ' + (i + 1) + ': ' + err.message + ' :: ' + err.stack);
    }
  }

  Logger.log('Finished. Processed ' + processedThisRun + ' row(s) this run.');
}

// -------- Core Processing --------
function createFolderCopyAndCallCentralized(propertyAddress, zipCode, rowIndex, lotSize, askingPrice) {
  var fn = 'createFolderCopyAndCallCentralized';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(THRESHOLD_SHEET_NAME);
  var sheetRow = rowIndex + 1; // convert 0-based array index to sheet row number
  var runId = Utilities.getUuid().slice(0, 8);

  Logger.log('[' + fn + ' ' + runId + '] Start address=' + propertyAddress + ' row=' + sheetRow); 
  try {
    // Create folder + copy
    var parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    var safeFolderName = sanitizeName_(propertyAddress);
    var newFolder = parentFolder.createFolder(safeFolderName);
    var templateFile = DriveApp.getFileById(TEMPLATE_FILE_ID);
    var newFileName = safeFolderName + ' - Analysis';
    var newFile = templateFile.makeCopy(newFileName, newFolder);
    var newFileUrl = newFile.getUrl();

    Logger.log('[' + runId + '] Created folder & copy: ' + newFileName); 

    // Permissions (service account + owner email if provided)
    var serviceAccountEmail = getProp_('SERVICE_ACCOUNT_EMAIL', '');
    if (serviceAccountEmail) { try { newFile.addEditor(serviceAccountEmail); } catch (e) { Logger.log('[' + runId + '] addEditor SA warn: ' + e.message); } }
    var ownerEmail = getProp_('OWNER_EMAIL', 'info@fortunatefoundations.com');
    if (ownerEmail) { try { newFile.addEditor(ownerEmail); } catch (e2) { Logger.log('[' + runId + '] addEditor owner warn: ' + e2.message); } }

    // Mark original sheet (checkbox + link) using 1-based columns
    sheet.getRange(sheetRow, COL.processedCheck + 1).setValue(true); // +1 convert 0-based index to 1-based col
    sheet.getRange(sheetRow, COL.link + 1).setValue(newFileUrl);

    // Populate copied spreadsheet
    populateNewSpreadsheet_(newFile.getId(), propertyAddress, zipCode, lotSize, askingPrice);

    // n8n webhook (optional)
    if (N8N_WEBHOOK_URL) {
      postJsonSafe_(N8N_WEBHOOK_URL, { fileName: newFileName, fileUrl: newFileUrl }, 'n8n', runId);
    }

    // Central script (optional) – allows the new file to set its own script properties after copy
    if (CENTRAL_URL) {
      var secret = Utilities.getUuid();
      postJsonSafe_(CENTRAL_URL, {
        action: 'initialize',
        spreadsheetId: newFile.getId(),
        callbackSecret: secret,
        propertyAddress: propertyAddress
      }, 'central', runId);
    }

    Logger.log('[' + runId + '] Done for ' + propertyAddress);
  } catch (err) {
    Logger.log('[' + fn + ' ' + runId + '] ERROR address=' + propertyAddress + ' row=' + sheetRow + ' :: ' + err.message);
    // Optionally flag error status in sheet if desired (not requested now)
  }
}

// -------- Helpers --------
function parseMetric_(raw) {
  if (raw == null || raw === '') return null;
  if (typeof raw === 'number') return raw > 1 ? raw / 100 : raw;
  var s = String(raw).trim();
  if (!s) return null;
  var pct = s.indexOf('%') !== -1;
  var num = parseFloat(s.replace(/[^0-9.+-]/g, ''));
  if (isNaN(num)) return null;
  return pct || num > 1 ? num / 100 : num; // treat >1 as percent form
}

function isChecked_(val) {
  return val === true || val === 1 || String(val).toUpperCase() === 'TRUE';
}

function sanitizeName_(name) {
  return String(name)
    .replace(/[\\/:*?"<>|]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 150);
}

function populateNewSpreadsheet_(fileId, address, zip, lotSize, askingPrice) {
  var ss = SpreadsheetApp.openById(fileId);
  var analysis = ss.getSheetByName('Detailed Analysis');
  var area = ss.getSheetByName('Area Summary');
  if (analysis) {
    analysis.getRange('B4').setValue(address + ', Austin, TX ' + (zip || ''));
    if (zip != null && String(zip).trim() !== '') analysis.getRange('B5').setValue(zip);
    analysis.getRange('B6').setValue(address);
    if (askingPrice != null && String(askingPrice).trim() !== '') analysis.getRange('B59').setValue(askingPrice);
  }
  if (area) {
    if (lotSize != null && String(lotSize).trim() !== '') area.getRange('B3').setValue(lotSize);
  }
  SpreadsheetApp.flush();
}

function postJsonSafe_(url, payloadObj, label, runId) {
  try {
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify(payloadObj)
    });
    var code = resp.getResponseCode();
    if (code >= 400) {
      Logger.log('[' + runId + '] ' + label + ' HTTP ' + code + ' body=' + resp.getContentText().slice(0, 250));
    } else {
      Logger.log('[' + runId + '] ' + label + ' HTTP ' + code);
    }
  } catch (e) {
    Logger.log('[' + runId + '] ' + label + ' fetch error: ' + e.message);
  }
}

// (Auto-sort removed from this file; keep separate sorter script.)