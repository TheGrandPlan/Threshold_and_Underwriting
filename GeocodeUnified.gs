// =============================================================
// GeocodeUnified.gs - Consolidated geocoding (replaces older variants when adopted)
//  - Batch geocodes rows in a sheet
//  - Resumable via Script Property
//  - Uses Script Properties for API key & state (NO hard-coded key)
//  - Designed for infrequent (monthly) runs
// =============================================================

// Configuration specific to the geocoding target sheet
var GEOCODE_CFG = {
  spreadsheetId: cfgProp('GEOCODE_SPREADSHEET_ID', ''), // set in Script Properties
  sheetName: cfgProp('GEOCODE_SHEET_NAME', 'Current Comps'),
  startRow: parseInt(cfgProp('GEOCODE_START_ROW', '2'), 10) || 2,
  col: { // 1-based indexes or letters as needed
    street: 'A',
    city: 'C',
    state: 'D',
    zip: 'E',
    output: 42 // AP
  },
  maxRuntimeMs: 5 * 60 * 1000, // 5 minutes (leave buffer for 6-min limit)
  perRunRowCap: parseInt(cfgProp('GEOCODE_PER_RUN_CAP', String(CFG.geocodeBatchSize)), 10) || 250,
  requestDelayMs: 50,
  maxRetries: 3
};

/** Entry point for batch geocoding */
function geocodeBatch() {
  var mod = 'Geocode';
  if (!CFG.geocodeKey) { log_(mod, 'API key missing (GEOCODE_API_KEY). Aborting.'); return; }
  if (!GEOCODE_CFG.spreadsheetId) { log_(mod, 'Spreadsheet ID missing (GEOCODE_SPREADSHEET_ID). Aborting.'); return; }
  var ss, sheet;
  try {
    ss = SpreadsheetApp.openById(GEOCODE_CFG.spreadsheetId);
    sheet = ss.getSheetByName(GEOCODE_CFG.sheetName);
  } catch (e) {
    log_(mod, 'Open error: ' + e.message); return;
  }
  if (!sheet) { log_(mod, 'Sheet not found: ' + GEOCODE_CFG.sheetName); return; }

  var props = PropertiesService.getScriptProperties();
  var lastProcessed = parseInt(props.getProperty(CFG.geocodeLastRowKey), 10);
  if (!lastProcessed || lastProcessed < GEOCODE_CFG.startRow - 1) lastProcessed = GEOCODE_CFG.startRow - 1;
  var lastDataRow = sheet.getLastRow();
  var startRow = lastProcessed + 1;
  if (startRow > lastDataRow) { log_(mod, 'Nothing to do (start beyond last row).'); return; }

  log_(mod, 'Starting at row ' + startRow + ' of ' + lastDataRow);
  var t0 = Date.now();
  var processed = 0;
  var apiCalls = 0;

  while (startRow <= lastDataRow && processed < GEOCODE_CFG.perRunRowCap) {
    if (Date.now() - t0 > GEOCODE_CFG.maxRuntimeMs) { log_(mod, 'Time buffer reached, stopping.'); break; }

    var street = sheet.getRange(GEOCODE_CFG.col.street + startRow).getValue();
    var city = sheet.getRange(GEOCODE_CFG.col.city + startRow).getValue();
    var state = sheet.getRange(GEOCODE_CFG.col.state + startRow).getValue();
    var zip = sheet.getRange(GEOCODE_CFG.col.zip + startRow).getValue();
    var outputCell = sheet.getRange(startRow, GEOCODE_CFG.col.output);
    var existing = outputCell.getValue();

    var addr = buildAddress_(street, city, state, zip);
    if (!addr.valid) {
      outputCell.setValue('Incomplete Address');
      lastProcessed = startRow; processed++; startRow++; Utilities.sleep(10); continue;
    }
    if (existing && String(existing).indexOf(',') !== -1 && String(existing).toUpperCase().indexOf('ERROR') !== 0) {
      lastProcessed = startRow; processed++; startRow++; continue; // already has coords
    }

    var attempt = 0; var success = false; var coord = '';
    while (attempt < GEOCODE_CFG.maxRetries && !success) {
      if (attempt > 0) {
        var wait = Math.pow(2, attempt) * 1000 + Math.floor(Math.random()*500);
        Utilities.sleep(wait);
      }
      try {
        var url = 'https://maps.googleapis.com/maps/api/geocode/json?address=' + encodeURIComponent(addr.full) + '&key=' + encodeURIComponent(CFG.geocodeKey);
        var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        apiCalls++;
        var code = resp.getResponseCode();
        if (code !== 200) throw new Error('HTTP ' + code);
        var data = JSON.parse(resp.getContentText());
        if (data.status === 'OK' && data.results && data.results.length) {
          var loc = data.results[0].geometry.location;
          coord = loc.lat + ',' + loc.lng;
          outputCell.setValue(coord);
          success = true;
        } else if (data.status === 'ZERO_RESULTS') {
          outputCell.setValue('Not Found');
          success = true; // definitive
        } else if (data.status === 'OVER_QUERY_LIMIT') {
          log_(mod, 'OVER_QUERY_LIMIT at row ' + startRow + ' attempt ' + attempt);
        } else {
          outputCell.setValue('Error:' + data.status);
          success = true; // treat as final for now
        }
      } catch (e) {
        log_(mod, 'Exception row ' + startRow + ': ' + e.message);
      }
      attempt++;
      if (!success && attempt < GEOCODE_CFG.maxRetries) Utilities.sleep(250);
    }

    if (!success) { log_(mod, 'Row ' + startRow + ' unresolved; will retry next run.'); break; }
    lastProcessed = startRow;
    processed++; startRow++;
    Utilities.sleep(GEOCODE_CFG.requestDelayMs);
  }

  props.setProperty(CFG.geocodeLastRowKey, String(lastProcessed));
  log_(mod, 'Run complete. Last processed row ' + lastProcessed + '. Rows this run ' + processed + '. API calls ' + apiCalls + '.');
}

function buildAddress_(street, city, state, zip) {
  var parts = [street, city, state, zip].map(function(p){ return (p==null)?'':String(p).trim(); });
  var valid = parts[0] !== '' && parts[1] !== '' && parts[2] !== '' && parts[3] !== '';
  return { full: parts.filter(Boolean).join(', '), valid: valid };
}
