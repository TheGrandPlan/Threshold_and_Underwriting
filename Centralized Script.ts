// Centralized Script (Fortunate Owned)
const AUTH_LOG_SPREADSHEET_ID = '1fEpV7dCBSwI4Tn45d3e1LsNJtO7hu9d3bgXXthmORGA'; 
const AUTH_LOG_SHEET_NAME = 'Authorizations'; 
// --- Standard Logging ---
const logSummary = [];
function log(message) {
  Logger.log(message); // Internal log for Centralized (visible in its own Executions if run directly)
  logSummary.push(message);
}
// --- End Logging ---

// === Web App Entry Points ===
function doPost(e) {
  let requestData;
  let responseService = ContentService.createTextOutput(); // Default to text
  responseService.setMimeType(ContentService.MimeType.TEXT); // Default MIME type

  try {
    requestData = JSON.parse(e.postData.contents);
    log(`📬 [Centralized] Received POST: Action = ${requestData.action}, SpreadsheetID = ${requestData.spreadsheetId}`);

    // --- Action Routing ---
    if (requestData.action === 'initialize') {
       if (!requestData.spreadsheetId || !requestData.callbackSecret) {
           throw new Error("Missing 'spreadsheetId' or 'callbackSecret' for initialize action.");
       }
       // handleSheetSetup now returns a detailed success message or throws a detailed error
       responseService.setContent(handleSheetSetup(requestData));

    } else if (requestData.action === 'getCredentials') {
        if (!requestData.spreadsheetId || !requestData.callbackSecret) {
           throw new Error("Missing 'spreadsheetId' or 'callbackSecret' for getCredentials action.");
       }
       // Returns JSON directly
       responseService = handleCredentialRequest(requestData); // Gets ContentService object back

    // *** ADDED: Route 'refreshCredentials' action ***
    } else if (requestData.action === 'refreshCredentials') {
         if (!requestData.spreadsheetId || !requestData.originalSecret) { // Check for correct params
           throw new Error("Missing 'spreadsheetId' or 'originalSecret' for refreshCredentials action.");
       }
       // Returns JSON directly
       responseService = handleRefreshCredentials(requestData); // Gets ContentService object back

    } else {
        throw new Error(`[Centralized] Unknown action: ${requestData.action}`);
    }
    // --- End Action Routing ---

  } catch (error) {
    log(`❌ [Centralized] Error processing POST request: ${error.message} - Stack: ${error.stack}`);
    // Return error message as text content
    responseService.setContent(`Centralized Error: ${error.message}`);
    // Ensure MIME type is TEXT for errors handled here
    responseService.setMimeType(ContentService.MimeType.TEXT);
  }

  return responseService;
}

function doGet() {
    // Keep existing doGet
    log("[Centralized] Received GET request.");
    return ContentService.createTextOutput("Centralized Script (with Auth Log) is active.")
                         .setMimeType(ContentService.MimeType.TEXT);
}
// --- End Web App Entry Points ---

// === Action Handlers ===

/**
 * Stores secret, sets description on the target sheet.
 * @param {object} data The request data {spreadsheetId, callbackSecret}
 * @return {string} Detailed success message.
 * @throws {Error} If any step fails (caught by doPost).
 */
function handleSheetSetup(data) {
    const { spreadsheetId, callbackSecret } = data;
    const cache = CacheService.getScriptCache();
    const CACHE_EXPIRATION_SECONDS = 1800; // 30 minutes
    const cacheKey = `secret_${spreadsheetId}`;
    let internalLog = []; // Log specific steps for the response message

    // --- Step 1: Store Secret ---
    try {
        cache.put(cacheKey, callbackSecret, CACHE_EXPIRATION_SECONDS);
        internalLog.push(`Stored secret for ${spreadsheetId} in cache.`);
        log(`🔑 ${internalLog[internalLog.length-1]}`); // Also log internally
    } catch (cacheErr) {
         log(`❌ Error storing secret in cache: ${cacheErr.message}`);
         throw new Error(`Failed to store secret in cache: ${cacheErr.message}`);
    }

    // --- Step 2: Get Web App URL ---
    let centralWebAppUrl;
    try {
        centralWebAppUrl = ScriptApp.getService().getUrl();
        if (!centralWebAppUrl) {
            throw new Error("ScriptApp.getService().getUrl() returned null or empty.");
        }
        internalLog.push(`Determined own URL.`);
        log(`🔗 ${internalLog[internalLog.length-1]}`);
    } catch(urlErr) {
        log(`❌ Error getting own web app URL: ${urlErr.message}`);
        try { cache.remove(cacheKey); } catch(e) {} // Attempt cleanup
        throw new Error(`Could not determine own web app URL: ${urlErr.message}`);
    }

    // --- Step 3: Set Description ---
    try {
        const descriptionData = { requestUrl: centralWebAppUrl, uniqueSecret: callbackSecret };
        const file = DriveApp.getFileById(spreadsheetId); // Requires Drive scope
        file.setDescription(JSON.stringify(descriptionData));
        internalLog.push(`Set description on ${spreadsheetId}.`);
        log(`📝 ${internalLog[internalLog.length-1]}`);
    } catch (descErr) {
        log(`❌ Failed to set description for ${spreadsheetId}: ${descErr.message}`);
        try { cache.remove(cacheKey); } catch(e) {} // Attempt cleanup
        // Include file ID in error message
        throw new Error(`Failed to set description for sheet ${spreadsheetId}: ${descErr.message}`);
    }

    // --- Success ---
    const successMsg = `Centralized SUCCESS: ${internalLog.join(' ')}`;
    log(`✅ ${successMsg}`); // Log final success internally
    // Return the detailed success message to be sent back to Threshold
    return successMsg;
}

/**
 * Handles initial setup request from Threshold script.
 * Stores secret temporarily in cache AND persistently in Log Sheet.
 * Sets description on the target sheet.
 * @param {object} data The request data {spreadsheetId, callbackSecret, propertyAddress? (Optional)}
 * @return {string} Detailed success message.
 * @throws {Error} If any step fails (caught by doPost).
 */
function handleSheetSetup(data) {
    const { spreadsheetId, callbackSecret, propertyAddress = "N/A" } = data; // Get address if passed
    const cache = CacheService.getScriptCache();
    const CACHE_EXPIRATION_SECONDS = 1800; // 30 minutes for cache
    const cacheKey = `secret_${spreadsheetId}`;
    let internalLog = [];

    // --- Step 1: Store Secret in Cache (for quick initial fetch) ---
    try {
        cache.put(cacheKey, callbackSecret, CACHE_EXPIRATION_SECONDS);
        internalLog.push(`Stored secret in cache (expires ${CACHE_EXPIRATION_SECONDS}s).`);
        log(`🔑 ${internalLog[internalLog.length-1]}`);
    } catch (cacheErr) {
         log(`❌ Error storing secret in cache: ${cacheErr.message}`);
         // Proceed even if cache fails, rely on persistent log
         internalLog.push(`Cache store failed: ${cacheErr.message}.`);
    }

    // *** ADDED: Step 1.5: Log Authorization Persistently ***
    try {
        const logSheet = SpreadsheetApp.openById(AUTH_LOG_SPREADSHEET_ID).getSheetByName(AUTH_LOG_SHEET_NAME);
        if (!logSheet) throw new Error(`Auth Log Sheet "${AUTH_LOG_SHEET_NAME}" not found.`);
        const timestamp = new Date();
        // Check if record already exists to avoid duplicates (optional)
        // const existingRecord = findRecordInLogSheet(spreadsheetId);
        // if (!existingRecord) { ... }
        logSheet.appendRow([
             timestamp,              // Col A: Timestamp
             spreadsheetId,          // Col B: SpreadsheetID
             callbackSecret,         // Col C: GeneratedSecret
             propertyAddress,        // Col D: PropertyAddress
             "Pending",              // Col E: Status
             "",                     // Col F: CredentialsIssuedTimestamp (blank initially)
             ""                      // Col G: LastError (blank initially)
        ]);
        internalLog.push(`Stored authorization record in persistent log.`);
        log(`💾 ${internalLog[internalLog.length-1]}`);
    } catch (logErr) {
        log(`❌ CRITICAL: Error writing to persistent Auth Log: ${logErr.message} - ${logErr.stack}`);
        // Decide if this should be fatal - probably should, as refresh depends on it
        throw new Error(`Failed to store authorization persistently: ${logErr.message}`);
    }
    // *** END ADDED STEP ***


    // --- Step 2: Get Web App URL (Keep existing) ---
    let centralWebAppUrl;
    try {
        centralWebAppUrl = ScriptApp.getService().getUrl();
        if (!centralWebAppUrl) { throw new Error("ScriptApp.getService().getUrl() returned null or empty."); }
        internalLog.push(`Determined own URL.`);
        log(`🔗 ${internalLog[internalLog.length-1]}`);
    } catch(urlErr) { /* ... existing error handling ... */ throw urlErr; }

    // --- Step 3: Set Description (Keep existing) ---
    try {
        const descriptionData = { requestUrl: centralWebAppUrl, uniqueSecret: callbackSecret };
        const file = DriveApp.getFileById(spreadsheetId);
        file.setDescription(JSON.stringify(descriptionData));
        internalLog.push(`Set description on ${spreadsheetId}.`);
        log(`📝 ${internalLog[internalLog.length-1]}`);
    } catch (descErr) { /* ... existing error handling ... */ throw descErr; }

    const successMsg = `Centralized SUCCESS: ${internalLog.join(' ')}`;
    log(`✅ ${successMsg}`);
    return successMsg; // Return success message for Threshold script
}


/**
 * Handles credential request from Local script (attempts cache first).
 * @param {object} data {spreadsheetId, callbackSecret}
 * @returns {ContentService.TextOutput} JSON response (credentials or error).
 */
function handleCredentialRequest(data) {
    const { spreadsheetId, callbackSecret } = data;
    const cache = CacheService.getScriptCache();
    const cacheKey = `secret_${spreadsheetId}`;
    let responsePayload = {};
    let foundInCache = false;

    log(`🔐 [Centralized] Received credential request for ${spreadsheetId} (Attempting Cache). Validating secret...`);

    if (!spreadsheetId || !callbackSecret) {
        log(`   ❌ Missing spreadsheetId or callbackSecret in credential request.`);
        responsePayload = { error: "Missing required parameters." };
    } else {
        const expectedSecret = cache.get(cacheKey);

        if (expectedSecret && expectedSecret === callbackSecret) {
            log(`   ✅ Cache Secret validation successful for ${spreadsheetId}.`);
            foundInCache = true;
            cache.remove(cacheKey); // Remove from cache after successful use
            log(`   🗑️ Removed secret from cache.`);

            responsePayload = getCredentialsPayload(); // Get credentials

            // *** ADDED: Update Log Sheet Status ***
            if (!responsePayload.error) {
                 updateLogSheetStatus(spreadsheetId, "Issued via Cache", new Date(), "");
            }
            // *** END ADDED STEP ***

        } else if (!expectedSecret) {
            log(`   ❌ Cache Secret validation failed: Secret expired or not found in cache for ${cacheKey}.`);
            // DO NOT return error yet, let Local script try the refresh method
            // We return a specific error recognizable by the Local script
             responsePayload = { error: "Invalid or expired request token.", reason: "cache_miss" };
        } else {
            log(`   ❌ Cache Secret validation failed: Provided secret mismatch for ${spreadsheetId}.`);
             // This is a definite error, secrets don't match
             responsePayload = { error: "Invalid request token.", reason: "cache_mismatch" };
              updateLogSheetStatus(spreadsheetId, "Error: Cache Mismatch", new Date(), "Secret mismatch on cache check");
        }
    }

    return ContentService.createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);
}

// *** ADDED: Handler for refreshCredentials action ***
/**
 * Handles credential request from Local script using persistent log.
 * @param {object} data {spreadsheetId, originalSecret}
 * @returns {ContentService.TextOutput} JSON response (credentials or error).
 */
function handleRefreshCredentials(data) {
    const { spreadsheetId, originalSecret } = data;
    let responsePayload = {};

    log(`🔄 [Centralized] Received REFRESH credential request for ${spreadsheetId}. Validating secret from log...`);

    if (!spreadsheetId || !originalSecret) {
        log(`   ❌ Missing spreadsheetId or originalSecret in refresh request.`);
        responsePayload = { error: "Missing required parameters for refresh." };
    } else {
        const record = findRecordInLogSheet(spreadsheetId);

        if (record && record.data && record.data.generatedSecret === originalSecret) {
            log(`   ✅ Persistent Log Secret validation successful for ${spreadsheetId} (Row ${record.rowNumber}).`);

            // Optional: Add checks here based on timestamp or status if desired
            // e.g., if (record.data.status === 'Revoked') { ... return error ... }

            responsePayload = getCredentialsPayload(); // Get credentials

            // Update Log Sheet Status
            if (!responsePayload.error) {
                 updateLogSheetStatus(spreadsheetId, "Issued via Refresh", new Date(), ""); // Update status on success
            }

        } else if (!record) {
            log(`   ❌ Persistent Log validation failed: Record not found for ${spreadsheetId}.`);
            responsePayload = { error: "Authorization record not found." };
        } else {
            // Record found, but secret didn't match
            log(`   ❌ Persistent Log validation failed: Provided secret mismatch for ${spreadsheetId} (Row ${record.rowNumber}).`);
            responsePayload = { error: "Invalid original secret." };
             updateLogSheetStatus(spreadsheetId, "Error: Refresh Mismatch", new Date(), "Secret mismatch on refresh check");
        }
    }
     return ContentService.createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);
}


// *** ADDED: Helper to retrieve actual credentials ***
/**
 * Retrieves credentials from PropertiesService and formats payload.
 * @returns {object} Payload object containing credentials or an error object.
 */
function getCredentialsPayload() {
  let payload = {};
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const privateKey = scriptProps.getProperty('PRIVATE_KEY');
    const apiKey = scriptProps.getProperty('GEOCODING_API_KEY'); // Geocoding key
    const staticMapsApiKey = scriptProps.getProperty('STATIC_MAP_API_KEY'); // Static Map key
    const userEmail = scriptProps.getProperty('AUTHORIZED_USER_EMAIL');
    const gcpProjectId = 'austin-market-project';
    const serviceAccountEmail = scriptProps.getProperty('SERVICE_ACCOUNT_EMAIL');

    // Check ALL required keys
    if (!privateKey || !apiKey || !staticMapsApiKey || !userEmail || !serviceAccountEmail) {
      log(` ❌ CRITICAL: Centralized script credential configuration incomplete...`);
      payload = { error: "Internal server error: Centralized configuration incomplete." };
    } else {
      payload = {
        privateKey: privateKey,
        apiKey: apiKey,                     // Geocoding API Key
        staticMapsApiKey: staticMapsApiKey, // Static Maps API Key
        userEmail: userEmail,
        gcpProjectId: gcpProjectId,
        serviceAccountEmail: serviceAccountEmail
      };
      log(` ✅ Prepared credentials payload (incl. Static Map Key)`);
    }
  } catch (propErr) {
    log(` ❌ ERROR fetching script properties: ${propErr.message}`);
    payload = { error: "Failed to fetch credentials from script properties." };
  }

  return payload;
}

// *** ADDED: Helper to find record in Log Sheet ***
/**
 * Finds a record in the Authorization Log Sheet based on Spreadsheet ID.
 * @param {string} spreadsheetId The ID of the spreadsheet to find.
 * @returns {object|null} Object { rowNumber: number, data: { ... } } or null if not found/error.
 */
function findRecordInLogSheet(spreadsheetId) {
    const functionName = "findRecordInLogSheet";
     if (!spreadsheetId) return null;
    try {
        const logSheet = SpreadsheetApp.openById(AUTH_LOG_SPREADSHEET_ID).getSheetByName(AUTH_LOG_SHEET_NAME);
        if (!logSheet) { log(`[${functionName}] Error: Auth Log Sheet "${AUTH_LOG_SHEET_NAME}" not found.`); return null;}

        // Assuming SpreadsheetID is in Column B (index 1 in 0-based array)
        const idColumnValues = logSheet.getRange(2, 2, logSheet.getLastRow() - 1, 1).getValues(); // Read Col B from row 2

        for (let i = 0; i < idColumnValues.length; i++) {
            if (idColumnValues[i][0] === spreadsheetId) {
                const rowNumber = i + 2; // +2 because we started reading from row 2 (index 0 = row 2)
                // Optionally fetch the whole row data if needed frequently
                const rowData = logSheet.getRange(rowNumber, 1, 1, 7).getValues()[0]; // A:G
                const data = {
                    timestamp: rowData[0],
                    spreadsheetId: rowData[1],
                    generatedSecret: rowData[2],
                    propertyAddress: rowData[3],
                    status: rowData[4],
                    credentialsIssuedTimestamp: rowData[5],
                    lastError: rowData[6]
                };
                log(`[${functionName}] Found record for ${spreadsheetId} at row ${rowNumber}.`);
                return { rowNumber: rowNumber, data: data };
            }
        }
        log(`[${functionName}] Record for ${spreadsheetId} not found in log.`);
        return null; // Not found
    } catch (error) {
        log(`[${functionName}] Error accessing Auth Log Sheet: ${error.message}`);
        return null;
    }
}


// *** ADDED: Helper to update status in Log Sheet ***
/**
 * Updates the Status and Timestamp columns for a given spreadsheetId in the Log sheet.
 * @param {string} spreadsheetId The ID of the spreadsheet record to update.
 * @param {string} status The new status string.
 * @param {Date} timestamp The timestamp for the update.
 * @param {string} errorMessage Optional error message to log.
 */
function updateLogSheetStatus(spreadsheetId, status, timestamp, errorMessage = "") {
     const functionName = "updateLogSheetStatus";
     if (!spreadsheetId) return;
     try {
         const record = findRecordInLogSheet(spreadsheetId);
         if (record && record.rowNumber) {
             const logSheet = SpreadsheetApp.openById(AUTH_LOG_SPREADSHEET_ID).getSheetByName(AUTH_LOG_SHEET_NAME);
             if (!logSheet) { log(`[${functionName}] Error: Auth Log Sheet not found.`); return; }
             // Update Status (Col E - index 5), Timestamp (Col F - index 6), Error (Col G - index 7)
             logSheet.getRange(record.rowNumber, 5, 1, 3).setValues([[status, timestamp, errorMessage]]);
             log(`[${functionName}] Updated status to "${status}" for ${spreadsheetId} at row ${record.rowNumber}.`);
         } else {
              log(`[${functionName}] Could not update status for ${spreadsheetId} - record not found.`);
         }
     } catch (error) {
         log(`[${functionName}] Error updating log sheet status for ${spreadsheetId}: ${error.message}`);
     }
}

// --- End Action Handlers ---