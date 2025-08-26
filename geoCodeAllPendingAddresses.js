// ========================================================================
// --- GEOCODING SCRIPT FOR 'Current Comps' SHEET ---
// ========================================================================

// --- SCRIPT CONFIGURATION CONSTANTS ---
// API key pulled from Script Properties (key name: API_KEY)
const GEOCODING_API_KEY = (function(){
    try { return PropertiesService.getScriptProperties().getProperty('API_KEY'); } catch(e){ return ''; }
})();
const SPREADSHEET_ID_FOR_GEOCODING = '1oIty2TBjKQSblo_W7H75Ctsah9un6vW2h1K_IJF4_0I'; // Spreadsheet containing "Current Comps"
const SHEET_NAME_FOR_GEOCODING = "Current Comps";                           // Sheet to geocode
const START_ROW_IN_SHEET = 2;          // The actual first row number with data (after any headers)
const ADDRESS_STREET_COL_LETTER = 'A'; // Column letter for Street Address
const ADDRESS_CITY_COL_LETTER = 'C';   // Column letter for City
const ADDRESS_STATE_COL_LETTER = 'D';  // Column letter for State
const ADDRESS_ZIP_COL_LETTER = 'E';    // Column letter for Zip
const OUTPUT_GEOCODE_COL_NUMBER = 42;  // Column AP (1-based index) where 'lat,lng' will be written

// --- Rate Limiting & Retry Configuration ---
const REQUEST_DELAY_MS = 50;          // Delay between API calls in milliseconds (e.g., 50ms = ~20 QPS)
const MAX_RETRIES_PER_ADDRESS = 3;    // Max times to retry geocoding a single address
const MAX_ROWS_PER_EXECUTION = 7500;  // Max rows to process in one script run (set high to try all)
const SCRIPT_PROPERTY_KEY_LAST_ROW = 'GEOCODE_LAST_ROW_PROCESSED_CURRENT_COMPS'; // Unique key for PropertiesService

// --- Execution Time Management ---
const MAX_SCRIPT_EXECUTION_TIME_MS = 5 * 60 * 1000; // 5 minutes (300,000 ms), leaves buffer for 6-min limit

/**
 * Geocodes addresses in the specified sheet that haven't been processed yet or errored.
 * Manages API rate limits, retries on errors, and saves progress.
 * Can be run multiple times to process a large sheet in chunks.
 */
function geocodeAllPendingAddresses() {
    const functionName = "geocodeAllPendingAddresses";
    Logger.log(`[${functionName}] Starting...`);

    const scriptProps = PropertiesService.getScriptProperties();
    let spreadsheet;
    let sheet;

    try {
        spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID_FOR_GEOCODING);
        sheet = spreadsheet.getSheetByName(SHEET_NAME_FOR_GEOCODING);
    } catch (e) {
        Logger.log(`[${functionName}] Error opening spreadsheet/sheet: ${e.message}. Exiting.`);
        return;
    }

    if (!sheet) {
        Logger.log(`[${functionName}] Sheet "${SHEET_NAME_FOR_GEOCODING}" not found. Exiting.`);
        return;
    }
    Logger.log(`[${functionName}] Sheet "${SHEET_NAME_FOR_GEOCODING}" found successfully.`);

    // --- Timer Setup ---
    const executionStartTime = new Date().getTime();
    // --- End Timer Setup ---

    let lastRowSuccessfullyProcessed = parseInt(scriptProps.getProperty(SCRIPT_PROPERTY_KEY_LAST_ROW)) || (START_ROW_IN_SHEET - 1);
    const absoluteLastDataRowInSheet = sheet.getLastRow();
    let currentProcessingRow = lastRowSuccessfullyProcessed + 1;

    Logger.log(`   Last successfully processed row: ${lastRowSuccessfullyProcessed}. Starting from row: ${currentProcessingRow}. Sheet data ends at: ${absoluteLastDataRowInSheet}`);

    let rowsProcessedThisExecution = 0;
    let apiCallsThisExecution = 0;

    // Main processing loop
    while (currentProcessingRow <= absoluteLastDataRowInSheet && rowsProcessedThisExecution < MAX_ROWS_PER_EXECUTION) {
        // Check Elapsed Time
        const currentTime = new Date().getTime();
        if ((currentTime - executionStartTime) >= MAX_SCRIPT_EXECUTION_TIME_MS) {
            Logger.log(`[${functionName}] Nearing execution time limit. Saving progress and exiting.`);
            break;
        }

        const street = sheet.getRange(`${ADDRESS_STREET_COL_LETTER}${currentProcessingRow}`).getValue();
        const city = sheet.getRange(`${ADDRESS_CITY_COL_LETTER}${currentProcessingRow}`).getValue();
        const state = sheet.getRange(`${ADDRESS_STATE_COL_LETTER}${currentProcessingRow}`).getValue();
        const zip = sheet.getRange(`${ADDRESS_ZIP_COL_LETTER}${currentProcessingRow}`).getValue();
        const existingOutputCell = sheet.getRange(currentProcessingRow, OUTPUT_GEOCODE_COL_NUMBER);
        const existingOutputValue = existingOutputCell.getValue();

        const fullAddress = `${street || ''}, ${city || ''}, ${state || ''} ${zip || ''}`
                            .replace(/undefined/gi, '').replace(/null/gi, '') // Handle null/undefined if they sneak in
                            .replace(/ ,/g,',').replace(/, ,/g,',').replace(/,\s*$/, '').trim();


        if (!street || String(street).trim() === "" || fullAddress.length < 5) { // Basic check for a valid address start
            Logger.log(`   Row ${currentProcessingRow}: Street address is empty or address too short ('${fullAddress}'). Marking as 'Incomplete Address'.`);
            existingOutputCell.setValue("Incomplete Address");
            lastRowSuccessfullyProcessed = currentProcessingRow; // This row is "handled"
            currentProcessingRow++;
            rowsProcessedThisExecution++;
            Utilities.sleep(10); // Tiny sleep even for non-API ops
            continue;
        }

        // Skip if already successfully geocoded (contains a comma and not an error message)
        if (existingOutputValue && String(existingOutputValue).includes(',') && !String(existingOutputValue).toUpperCase().startsWith("ERROR")) {
            // Logger.log(`   Row ${currentProcessingRow}: Already geocoded for '${fullAddress}'. Value: '${existingOutputValue}'. Skipping.`);
            lastRowSuccessfullyProcessed = currentProcessingRow;
            currentProcessingRow++;
            rowsProcessedThisExecution++;
            continue;
        }

        Logger.log(`   Processing Row ${currentProcessingRow}: '${fullAddress}'`);
        let coordinates = null;
        let retries = 0;
        let geocodeSuccess = false; // Flag for successful geocoding attempt this iteration

        while (retries < MAX_RETRIES_PER_ADDRESS && !geocodeSuccess) {
            if (retries > 0) {
                const waitTime = Math.pow(2, retries) * 1000 + Math.floor(Math.random() * 1000); // Exponential backoff + jitter
                Logger.log(`      Retrying (${retries}/${MAX_RETRIES_PER_ADDRESS}) after ${(waitTime/1000).toFixed(1)}s...`);
                Utilities.sleep(waitTime);
            }
            try {
                if (!GEOCODING_API_KEY) {
                    Logger.log("      Missing API_KEY in Script Properties; aborting this run.");
                    break; // break retry loop; will exit main loop after marking
                }
                const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(fullAddress)}&key=${encodeURIComponent(GEOCODING_API_KEY)}`;
                const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
                apiCallsThisExecution++;
                const responseCode = response.getResponseCode();
                const responseBody = response.getContentText();
                const data = JSON.parse(responseBody);

                if (responseCode === 200 && data.status === "OK" && data.results && data.results.length > 0) {
                    const location = data.results[0].geometry.location;
                    coordinates = `${location.lat},${location.lng}`; // No space after comma
                    existingOutputCell.setValue(coordinates);
                    Logger.log(`      Success: ${coordinates}`);
                    geocodeSuccess = true;
                } else if (data.status === "OVER_QUERY_LIMIT") {
                    Logger.log(`      OVER_QUERY_LIMIT for '${fullAddress}'. Will retry (attempt ${retries + 1}).`);
                    retries++;
                    if (retries >= MAX_RETRIES_PER_ADDRESS) { // If max retries for OQL reached, note it
                         Logger.log(`      Max retries for OVER_QUERY_LIMIT reached for '${fullAddress}'.`);
                         existingOutputCell.setValue("ERROR: OVER_QUERY_LIMIT");
                    }
                } else if (data.status === "ZERO_RESULTS") {
                     Logger.log(`      ZERO_RESULTS for '${fullAddress}'.`);
                     existingOutputCell.setValue("Zero Results");
                     geocodeSuccess = true; // Mark as success to stop retrying this address
                } else {
                    Logger.log(`      API Error for '${fullAddress}': Code ${responseCode}, Status ${data.status}, Msg: ${data.error_message || responseBody.substring(0,150)}`);
                    existingOutputCell.setValue(`ERROR: ${data.status}`);
                    geocodeSuccess = true; // Mark as handled (error noted) to stop retrying
                }
            } catch (e) {
                Logger.log(`      Exception for '${fullAddress}': ${e.message}`);
                retries++;
                 if (retries >= MAX_RETRIES_PER_ADDRESS) {
                     existingOutputCell.setValue(`ERROR: Exception`);
                     Logger.log(`      Max retries for Exception reached for '${fullAddress}'.`);
                 }
            }
        } // End retry while loop

        if (geocodeSuccess) {
            lastRowSuccessfullyProcessed = currentProcessingRow; // Update only if geocoding attempt was conclusive (success or specific error)
        } else {
             // This means it failed all retries, likely due to persistent OVER_QUERY_LIMIT or repeated exceptions
             Logger.log(`   Failed to definitively geocode '${fullAddress}' after ${MAX_RETRIES_PER_ADDRESS} retries. Will retry this row on next full execution.`);
             // Do not update lastRowSuccessfullyProcessed, so it retries this row next time the script runs.
             break; // Exit the main processing loop for *this execution* to save progress.
        }

        currentProcessingRow++;
        rowsProcessedThisExecution++;
        Utilities.sleep(REQUEST_DELAY_MS); // Inter-request delay
    } // End main while loop

    scriptProps.setProperty(SCRIPT_PROPERTY_KEY_LAST_ROW, String(lastRowSuccessfullyProcessed));
    Logger.log(`[${functionName}] Execution finished. Last successfully processed row: ${lastRowSuccessfullyProcessed}. Processed ${rowsProcessedThisExecution} addresses this run. API calls: ${apiCallsThisExecution}.`);

    if (currentProcessingRow <= absoluteLastDataRowInSheet && rowsProcessedThisExecution >= MAX_ROWS_PER_EXECUTION) {
        Logger.log("   MAX_ROWS_PER_EXECUTION reached. More rows may need processing. Run script again.");
    } else if (currentProcessingRow <= absoluteLastDataRowInSheet) {
        Logger.log("   Script stopped (likely due to time limit or persistent error). More rows to process. Run script again.");
    } else if (lastRowSuccessfullyProcessed >= absoluteLastDataRowInSheet) {
        Logger.log("   All rows in the sheet appear to have been processed.");
        // Optional: Clear the property if you want it to start from scratch next time,
        // or if you want a clear signal that it completed everything.
        // scriptProps.deleteProperty(SCRIPT_PROPERTY_KEY_LAST_ROW);
        // Logger.log("   Cleared 'last processed row' property as all rows seem complete.");
    }
}