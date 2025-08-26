// ========================================================================
// GeocodeAllPendingAddresses (Production Copy)
//  - Copy this file into Apps Script for monthly geocoding runs.
//  - Uses Script Property 'API_KEY' (Google Geocoding API Key) – not hard-coded.
//  - Resumable via Script Property 'GEOCODE_LAST_ROW_PROCESSED_CURRENT_COMPS'.
// ========================================================================

// --- SCRIPT CONFIGURATION CONSTANTS ---
const GEOCODING_API_KEY = (function(){ try { return PropertiesService.getScriptProperties().getProperty('API_KEY'); } catch(e){ return ''; } })();
const SPREADSHEET_ID_FOR_GEOCODING = '1oIty2TBjKQSblo_W7H75Ctsah9un6vW2h1K_IJF4_0I'; // Spreadsheet containing "Current Comps"
const SHEET_NAME_FOR_GEOCODING = 'Current Comps';
const START_ROW_IN_SHEET = 2;          // First data row (after headers)
const ADDRESS_STREET_COL_LETTER = 'A'; // Street
const ADDRESS_CITY_COL_LETTER   = 'C'; // City
const ADDRESS_STATE_COL_LETTER  = 'D'; // State
const ADDRESS_ZIP_COL_LETTER    = 'E'; // Zip
const OUTPUT_GEOCODE_COL_NUMBER = 42;  // Column AP (1-based) for 'lat,lng' or status

// --- Rate Limiting & Retry Configuration ---
const REQUEST_DELAY_MS = 50;            // ~20 QPS theoretical
const MAX_RETRIES_PER_ADDRESS = 3;      // Retries for OVER_QUERY_LIMIT / transient errors
const MAX_ROWS_PER_EXECUTION = 7500;    // Safety upper bound (practically limited by time)
const SCRIPT_PROPERTY_KEY_LAST_ROW = 'GEOCODE_LAST_ROW_PROCESSED_CURRENT_COMPS';

// --- Execution Time Management ---
const MAX_SCRIPT_EXECUTION_TIME_MS = 5 * 60 * 1000; // 5 minutes, leave buffer for 6‑min simple trigger

/**
 * Public entry point: geocodes pending addresses batch-wise.
 */
function geocodeAllPendingAddresses() {
	const fn = 'geocodeAllPendingAddresses';
	Logger.log(`[${fn}] Starting...`);

	if (!GEOCODING_API_KEY) {
		Logger.log(`[${fn}] Missing API_KEY Script Property. Abort.`);
		return;
	}

	const props = PropertiesService.getScriptProperties();
	let sheet;
	try {
		const ss = SpreadsheetApp.openById(SPREADSHEET_ID_FOR_GEOCODING);
		sheet = ss.getSheetByName(SHEET_NAME_FOR_GEOCODING);
	} catch (e) {
		Logger.log(`[${fn}] Error opening spreadsheet/sheet: ${e.message}`);
		return;
	}
	if (!sheet) { Logger.log(`[${fn}] Sheet '${SHEET_NAME_FOR_GEOCODING}' not found.`); return; }
	Logger.log(`[${fn}] Sheet '${SHEET_NAME_FOR_GEOCODING}' opened.`);

	const startTime = Date.now();
	let lastRowDone = parseInt(props.getProperty(SCRIPT_PROPERTY_KEY_LAST_ROW), 10) || (START_ROW_IN_SHEET - 1);
	const lastDataRow = sheet.getLastRow();
	let row = lastRowDone + 1;
	let processed = 0;
	let apiCalls = 0;

	Logger.log(`   Resuming at row ${row}; sheet ends at ${lastDataRow}.`);

	while (row <= lastDataRow && processed < MAX_ROWS_PER_EXECUTION) {
		if (Date.now() - startTime >= MAX_SCRIPT_EXECUTION_TIME_MS) {
			Logger.log(`[${fn}] Time buffer reached; stopping early.`);
			break;
		}

		const street = sheet.getRange(`${ADDRESS_STREET_COL_LETTER}${row}`).getValue();
		const city   = sheet.getRange(`${ADDRESS_CITY_COL_LETTER}${row}`).getValue();
		const state  = sheet.getRange(`${ADDRESS_STATE_COL_LETTER}${row}`).getValue();
		const zip    = sheet.getRange(`${ADDRESS_ZIP_COL_LETTER}${row}`).getValue();
		const outCell = sheet.getRange(row, OUTPUT_GEOCODE_COL_NUMBER);
		const existing = outCell.getValue();

		const fullAddress = `${street || ''}, ${city || ''}, ${state || ''} ${zip || ''}`
			.replace(/undefined|null/gi,'')
			.replace(/ ,/g,',')
			.replace(/, ,/g,',')
			.replace(/,\s*$/,'')
			.trim();

		if (!street || String(street).trim() === '' || fullAddress.length < 5) {
			outCell.setValue('Incomplete Address');
			lastRowDone = row; row++; processed++; Utilities.sleep(10); continue;
		}

		if (existing && String(existing).includes(',') && !String(existing).toUpperCase().startsWith('ERROR')) {
			lastRowDone = row; row++; processed++; continue; // already done
		}

		Logger.log(`   Row ${row}: '${fullAddress}'`);
		let retries = 0; let success = false; let coordinates = null;
		while (retries < MAX_RETRIES_PER_ADDRESS && !success) {
			if (retries > 0) {
				const wait = Math.pow(2, retries) * 1000 + Math.floor(Math.random()*1000); // backoff + jitter
				Logger.log(`      Retry ${retries}/${MAX_RETRIES_PER_ADDRESS} after ${(wait/1000).toFixed(1)}s`);
				Utilities.sleep(wait);
			}
			try {
				const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(fullAddress)}&key=${encodeURIComponent(GEOCODING_API_KEY)}`;
				const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
				apiCalls++;
				const code = response.getResponseCode();
				const body = response.getContentText();
				const data = JSON.parse(body);
				if (code === 200 && data.status === 'OK' && data.results && data.results.length > 0) {
					const loc = data.results[0].geometry.location;
						coordinates = `${loc.lat},${loc.lng}`;
						outCell.setValue(coordinates);
						success = true;
						Logger.log(`      Success: ${coordinates}`);
				} else if (data.status === 'OVER_QUERY_LIMIT') {
					Logger.log(`      OVER_QUERY_LIMIT hit (attempt ${retries + 1}).`);
					retries++;
					if (retries >= MAX_RETRIES_PER_ADDRESS) {
						outCell.setValue('ERROR: OVER_QUERY_LIMIT');
					}
				} else if (data.status === 'ZERO_RESULTS') {
					outCell.setValue('Zero Results');
					success = true;
				} else {
					outCell.setValue(`ERROR: ${data.status}`);
					Logger.log(`      API Error status ${data.status}`);
					success = true; // treat as handled
				}
			} catch (e) {
				Logger.log(`      Exception: ${e.message}`);
				retries++;
				if (retries >= MAX_RETRIES_PER_ADDRESS) {
					outCell.setValue('ERROR: Exception');
				}
			}
		}

		if (success) {
			lastRowDone = row;
		} else {
			Logger.log(`   Row ${row} unresolved; will retry next run.`);
			break; // do not advance lastRowDone so we retry the same row next execution
		}

		row++; processed++;
		Utilities.sleep(REQUEST_DELAY_MS);
	}

	PropertiesService.getScriptProperties().setProperty(SCRIPT_PROPERTY_KEY_LAST_ROW, String(lastRowDone));
	Logger.log(`[${fn}] Finished. Last processed row: ${lastRowDone}. Rows this run: ${processed}. API calls: ${apiCalls}.`);
	if (row <= lastDataRow) {
		Logger.log('   More rows remain – run again to continue.');
	} else {
		Logger.log('   All rows processed.');
	}
}

