// ==================================================
// Local Script Template (geoComp) - Minimal Cleanup
//  - Uses Script Properties for IDs (PRELIMINARY_SHEET_ID, DATA_SPREADSHEET_ID, SLIDES_TEMPLATE_ID)
//  - Optional DEBUG logging via Script Property: DEBUG_LOG = 'true'
//  - Removed duplicate runInvestorSplitOptimization definition
//  - Increased setup trigger delay to ~60s for reliability
// ==================================================

/**
 * @OnlyCurrentDoc Limits the script to only accessing the current spreadsheet.
 */

function getProp(key, fallback) { try { var v = PropertiesService.getScriptProperties().getProperty(key); return (v !== null && v !== '' ? v : fallback); } catch(e){ return fallback; } }
const PRELIMINARY_SHEET_ID = getProp('PRELIMINARY_SHEET_ID','1oIty2TBjKQSblo_W7H75Ctsah9un6vW2h1K_IJF4_0I');
const PRELIMINARY_SHEET_NAME = 'Deal Analysis Summary - Prospective Properties';
const SLIDES_TEMPLATE_ID = getProp('SLIDES_TEMPLATE_ID','15E264Ht5CXUxtv8y0Ky43zrp9GUHJZxS4vzmDhWDnpY'); 
const CHART_SHEET_NAME = 'Executive Summary'; // <<< e.g., 'Executive Summary' or 'Detailed Analysis'
const PIE_CHART_1_TITLE = 'Project Timeline'; // <<< CASE-SENSITIVE
const PIE_CHART_2_TITLE = 'Investment vs. Profit';
//const TARGET_INVESTOR_IRR = 0.50; // <<< TARGET: Set desired IRR (e.g., 25%)
const TARGET_IRR_INPUT_CELL = 'D153';
const INVESTOR_SPLIT_CELL = 'B145'; // Cell to CHANGE
const INVESTOR_IRR_CELL = 'B153';   // Cell to READ 
const PROJECT_NET_PROFIT_CELL = 'B137'; // Cell containing overall project net profit

const context = {
    spreadsheet: null,
    sheet: null,
    dataSpreadsheet: null,
    dataSheet: null,
    compProperties: [],
    config: {
        MASTER_SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
        MASTER_SHEET_NAME: 'Sales Comps',
        ADDRESS_CELL: 'A3',
        COMP_RADIUS_CELL: 'P4',
        DATE_FILTER_CELL: 'P5',
        AGE_FILTER_CELL: 'P7',
        SIZE_FILTER_CELL: 'P9',
        SUBJECT_SIZE_CELL: 'G3',
        ANNUNCIATOR_CELL: 'P10',
        SD_MULTIPLIER_CELL: 'P11',
        COMP_RESULTS_START_ROW: 33,
        COMP_RESULTS_START_COLUMN: 'A',
    DATA_SPREADSHEET_ID: getProp('DATA_SPREADSHEET_ID','1oIty2TBjKQSblo_W7H75Ctsah9un6vW2h1K_IJF4_0I'),
        DATA_SHEET_NAME: "Current Comps",
    },
    filters: {
        radius: 0,
        date: null,
        yearBuilt: null,
        sizePercentage: 0
    },
    visibleRows: []
};

function initializeContext() {
    // Initialize the main spreadsheet if not already done
    if (!context.spreadsheet) {
        try {
            context.spreadsheet = SpreadsheetApp.openById(context.config.MASTER_SPREADSHEET_ID);
        } catch (error) {
            Logger.log(`Error opening master spreadsheet: ${error}`);
            return; // Exit if the master spreadsheet cannot be opened
        }
    }

    // Initialize the main sheet if not already done
    if (!context.sheet) {
        context.sheet = context.spreadsheet.getSheetByName(context.config.MASTER_SHEET_NAME);
        if (!context.sheet) {
            Logger.log(`Sheet "${context.config.MASTER_SHEET_NAME}" not found.`);
            return; // Exit if the master sheet cannot be found
        }
    }

    // Initialize the data spreadsheet if not already done
    if (!context.dataSpreadsheet) {
        try {
            context.dataSpreadsheet = SpreadsheetApp.openById(context.config.DATA_SPREADSHEET_ID);
        } catch (error) {
            Logger.log(`Error opening data spreadsheet: ${error}`);
            return; // Exit if the data spreadsheet cannot be opened
        }
    }

    // Initialize the data sheet if not already done
    if (!context.dataSheet) {
        context.dataSheet = context.dataSpreadsheet.getSheetByName(context.config.DATA_SHEET_NAME);
        if (!context.dataSheet) {
            Logger.log(`Sheet "${context.config.DATA_SHEET_NAME}" not found.`);
            return; // Exit if the data sheet cannot be found
        }
    }
    // Log the complete context object to debug issues
    Logger.log("Context initialized: " + JSON.stringify(context));
}


// === Setup Functions ===

/**
 * Simple trigger that runs when the spreadsheet is opened.
 * Adds a custom menu to manually start the setup if not complete.
 * @param {object} e The event object.
 */
function onOpen(e) {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('⚙️ Setup'); // Create a custom menu
    try {
        const scriptProps = PropertiesService.getScriptProperties();
        const setupComplete = scriptProps.getProperty('SETUP_COMPLETE');

        if (setupComplete !== 'true') {
            Logger.log("onOpen: Setup not complete. Adding setup menu item.");
            menu.addItem('▶️ Run Initial Setup (Required Once)', 'runInitialSetup'); // Add menu item to run the function
            SpreadsheetApp.getActiveSpreadsheet().toast('Please run "⚙️ Setup > ▶️ Run Initial Setup" once to initialize.', 'Setup Required', 15);
        } else {
            Logger.log("onOpen: Setup already complete.");
            menu.addItem('Setup Complete', 'setupAlreadyDone'); // Indicate completion (optional)
        }
    } catch (err) {
         Logger.log(`onOpen: Error creating menu: ${err.message}`);
         ui.alert(`Error setting up menu: ${err.message}`);
         menu.addItem('Error creating setup menu', 'setupAlreadyDone'); // Add error indicator
    }
    menu.addToUi(); // Add the menu to the spreadsheet UI
    // --- Slides Menu ---
    ui.createMenu('📊 Slides')
        .addItem('▶️ Generate Presentation', 'createPresentationFromSheet')
        .addToUi();
        // --- Calculations Menu ---
    ui.createMenu('⚙️ Calculations') // New menu
        .addItem('Optimize Investor Split', 'runInvestorSplitOptimization') // Points to a wrapper function
        .addSeparator() // Optional separator
        .addItem('Execute Break-Even Analysis', 'runBreakevenAnalysis') // New item
        .addItem('Reset Break-Even Inputs', 'resetBreakevenInputs')   // New item
        .addToUi();
}
// Dummy function for menu item when setup is done or errored
function setupAlreadyDone() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Setup has already been completed or encountered an error during menu creation.', 'Info', 5);
}

// --- Menu Item Wrapper Function ---
/**
 * Reads target IRR from sheet, validates it, then calls the
 * goal seek calculation function. Triggered by menu item.
 */
function runInvestorSplitOptimization() {
    const functionName = "runInvestorSplitOptimization";
    const ui = SpreadsheetApp.getUi();
    Logger.log(`[${functionName}] Manual trigger: Running Investor Split Optimization...`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailedAnalysisSheet = ss.getSheetByName('Detailed Analysis');

    if (!detailedAnalysisSheet) {
         ui.alert("Sheet 'Detailed Analysis' not found.");
         Logger.log(`[${functionName}] Optimization failed: Detailed Analysis sheet not found.`);
         return;
    }

    try {
        // --- Read Target IRR from the Sheet Cell ---
        const targetIRRCell = detailedAnalysisSheet.getRange(TARGET_IRR_INPUT_CELL);
        const targetIRRValue = targetIRRCell.getValue();
        Logger.log(`   Read value from ${TARGET_IRR_INPUT_CELL}: ${targetIRRValue} (Type: ${typeof targetIRRValue})`);

        // --- Validate Input ---
        if (typeof targetIRRValue !== 'number' || isNaN(targetIRRValue)) {
             throw new Error(`Invalid Target IRR in cell ${TARGET_IRR_INPUT_CELL}. Please enter a valid percentage (e.g., 25 for 25%).`);
        }
        // Ensure it's within a reasonable range (e.g., 0% to 500% stored as 0 to 5)
        if (targetIRRValue <= 0 || targetIRRValue > 5) { // Value stored as decimal (25% = 0.25)
            throw new Error(`Target IRR (${targetIRRCell.getDisplayValue()}) in cell ${TARGET_IRR_INPUT_CELL} is outside the expected range (e.g., 1% to 500%).`);
        }
        // The value read is already the decimal (e.g., 0.25 for 25%) because the cell is formatted as %
        const targetIRRDecimal = targetIRRValue;
        const targetIRRDisplay = targetIRRCell.getDisplayValue(); // Get the formatted string for messages

        Logger.log(`   Using Target IRR: ${targetIRRDisplay} (${targetIRRDecimal})`);

        // --- Call the Goal Seek Function ---
        const finalSplit = calculateInvestorSplitForTargetIRR(
            detailedAnalysisSheet,
            targetIRRDecimal, // Pass the validated decimal value
            INVESTOR_SPLIT_CELL,
            INVESTOR_IRR_CELL,
            PROJECT_NET_PROFIT_CELL
        );

        // --- Report Result ---
        if (finalSplit !== null) {
            const finalSplitDisplay = (finalSplit * 100).toFixed(2);
            // Read the actual final IRR achieved
            const finalAchievedIRR = detailedAnalysisSheet.getRange(INVESTOR_IRR_CELL).getDisplayValue();
            ui.alert(`Optimization complete.\nTarget IRR: ${targetIRRDisplay}\nAchieved IRR: ${finalAchievedIRR}\n\nInvestor split set to approx: ${finalSplitDisplay}%`);
            Logger.log(`   Optimization successful. Final Split: ${finalSplitDisplay}%, Achieved IRR: ${finalAchievedIRR}`);
        } else {
            ui.alert(`Optimization could not complete successfully (Target: ${targetIRRDisplay}). Check logs or project profitability.`);
            Logger.log(`   Optimization did not complete successfully.`);
        }

    } catch (error) {
         Logger.log(`[${functionName}] Error: ${error.message}`);
         ui.alert(`Error during optimization: ${error.message}`);
    }
    Logger.log(`[${functionName}] Manual trigger: Investor Split Optimization finished.`);
}
/**
 * Function triggered MANUALLY the first time via the '⚙️ Setup' menu,
 * OR later automatically by its own time trigger (this time trigger is deleted after successful setup).
 * Orchestrates the entire one-time setup process: fetches credentials, stores them,
 * deletes the setup time trigger, creates the necessary installable onEdit trigger,
 * marks setup complete, and runs the initial analysis.
 */
function runInitialSetup() {
    const scriptProps = PropertiesService.getScriptProperties();
    let setupComplete = scriptProps.getProperty('SETUP_COMPLETE'); // Get current status
    const functionName = "runInitialSetup"; // For logging

    // --- Declare variables at function scope ---
    let ui = null;
    let isManualRun = false;
    let credentials = null;
    // --- End variable declaration ---

    // --- Determine Execution Context & Get UI Conditionally ---
    try {
        if (Session.getActiveUser() != null) {
             ui = SpreadsheetApp.getUi();
             isManualRun = true;
        }
    } catch (e) {
         Logger.log(`[${functionName}] Could not get UI (Error: ${e.message}). Assuming non-interactive execution.`);
         isManualRun = false;
    }
    // --- End Context Determination ---

    // --- Step 0: Check if setup is already done ---
    // Check for 'true' or 'trigger_error' states which mean setup attempted/finished
    if (setupComplete === 'true' || setupComplete === 'trigger_error') {
        Logger.log(`[${functionName}] 🏁 Setup already attempted/completed (Status: ${setupComplete}).`);
        if (isManualRun && ui) {
             ui.alert(`Setup has already been completed or encountered an error previously (Status: ${setupComplete}).`);
        }
        // Ensure the time trigger is gone even if setup finished with trigger error
        deleteOwnSetupTrigger();
        return;
    }
    // --- End Step 0 ---

    // --- Step 0.5: Ensure Setup Time Trigger Exists or Create It ---
    // This logic primarily handles the very first manual run to bootstrap the automated process
    if (!doesSetupTriggerExist()) {
        if (!isManualRun) {
            Logger.log(`[${functionName}] ❌ ERROR: Setup time trigger missing, but script is not running in a user context. Cannot create trigger. Please run setup manually via the menu first.`);
            // Set status to error maybe?
            // scriptProps.setProperty('SETUP_COMPLETE', 'error_no_trigger');
            return;
        }
        Logger.log(`[${functionName}] 🚦 Trigger check: No setup time trigger found. Creating one now (requires user authorization via menu click)...`);
        try {
            createSetupTrigger(); // Creates the time-based trigger to run this function again soon
            Logger.log(`[${functionName}]    ⏲️ Setup time trigger created successfully by user action.`);
            if (ui) {
                 ui.alert('Setup time trigger created. The automated setup process will start in approximately 1-5 minutes. You can close this sheet.');
            }
            return; // Exit after creating trigger; the triggered run will perform the actual setup
        } catch (err) {
            Logger.log(`[${functionName}]    ❌ Failed to create setup time trigger during manual run: ${err.message} - Stack: ${err.stack || 'N/A'}`);
            if (ui) {
                ui.alert(`Error creating the setup time trigger: ${err.message}. Please check script permissions (script.scriptapp scope) and try running setup again.`);
            }
            return; // Stop if trigger creation failed
        }
    } else {
         Logger.log(`[${functionName}] 🚦 Trigger check: Setup time trigger already exists. Proceeding with automated setup logic (likely running via trigger)...`);
    }
    // --- End Step 0.5 ---

    Logger.log(`[${functionName}] 🚀🚀🚀 Running Initial Setup Steps... 🚀🚀🚀`);
    let overallSuccess = false; // Track if all steps succeed

    try {
        // --- Step 1: Read Description ---
        Logger.log("[${functionName}]    1. Reading spreadsheet description...");
        const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
        const file = DriveApp.getFileById(ssId); // Requires drive scope
        const description = file.getDescription();
        if (!description) throw new Error("Spreadsheet description is empty or missing.");
        let descriptionData;
        try { descriptionData = JSON.parse(description); } catch (e) { throw new Error(`Failed to parse description JSON: ${e.message}`); }
        const { requestUrl, uniqueSecret } = descriptionData;
        if (!requestUrl || !uniqueSecret) throw new Error("Missing 'requestUrl' or 'uniqueSecret' in description metadata.");
        Logger.log(`[${functionName}]       Found Request URL: ${requestUrl}`);
        Logger.log(`[${functionName}]       Found Secret: ${uniqueSecret.substring(0,8)}...`);
        // --- End Step 1 ---

        // --- Step 2: Call Centralized for Credentials ---
        Logger.log("[${functionName}]    2. Fetching credentials from Centralized...");
        let credentialsFetched = false;
        let centralResponse;
        let centralResponseBody;
        let centralStatusCode;

        // --- Attempt 1: Get from Cache via 'getCredentials' action ---
        try {
            Logger.log(`[${functionName}]       Attempt 1: Using 'getCredentials' (cache)...`);
            centralResponse = UrlFetchApp.fetch(requestUrl, {
                method: 'post', contentType: 'application/json',
                payload: JSON.stringify({ action: 'getCredentials', spreadsheetId: ssId, callbackSecret: uniqueSecret }),
                muteHttpExceptions: true
            });
            centralStatusCode = centralResponse.getResponseCode();
            centralResponseBody = centralResponse.getContentText();
            Logger.log(`[${functionName}]       'getCredentials' Response - Status: ${centralStatusCode}, Body: ${centralResponseBody.substring(0, 100)}...`);

            if (centralStatusCode === 200) {
                credentials = JSON.parse(centralResponseBody);
                if (credentials && !credentials.error) {
                    Logger.log(`[${functionName}]       ✅ Credentials received successfully via cache.`);
                    credentialsFetched = true;
                } else if (credentials && credentials.error && credentials.reason === 'cache_miss') {
                     Logger.log(`[${functionName}]       ⚠️ Cache miss reported by Centralized. Will attempt refresh.`);
                     // Do not throw error yet, proceed to refresh attempt
                } else if (credentials && credentials.error) {
                     // Other definite error (like mismatch)
                     throw new Error(`Centralized returned an application error: ${credentials.error}`);
                } else {
                     // Unexpected successful response format
                     throw new Error(`Centralized returned unexpected success payload: ${centralResponseBody}`);
                }
            } else {
                // HTTP error during 'getCredentials' call
                 throw new Error(`Failed to fetch credentials (cache attempt). Centralized returned HTTP ${centralStatusCode}. Body: ${centralResponseBody}`);
            }
        } catch (e) {
             Logger.log(`[${functionName}]       ❌ Error during 'getCredentials' attempt: ${e.message}`);
             // If JSON parsing failed, throw error
             if (e.message.includes("JSON")) { throw e; }
             // Otherwise, allow proceeding to refresh attempt if it wasn't a specific known error yet
        }


        // --- Attempt 2: Refresh from Log via 'refreshCredentials' action (if needed) ---
        if (!credentialsFetched) {
            Logger.log(`[${functionName}]       Attempt 2: Using 'refreshCredentials' (persistent log)...`);
            try {
                centralResponse = UrlFetchApp.fetch(requestUrl, {
                    method: 'post', contentType: 'application/json',
                    payload: JSON.stringify({ action: 'refreshCredentials', spreadsheetId: ssId, originalSecret: uniqueSecret }), // Use originalSecret here
                    muteHttpExceptions: true
                });
                centralStatusCode = centralResponse.getResponseCode();
                centralResponseBody = centralResponse.getContentText();
                 Logger.log(`[${functionName}]       'refreshCredentials' Response - Status: ${centralStatusCode}, Body: ${centralResponseBody.substring(0, 100)}...`);

                 if (centralStatusCode === 200) {
                     credentials = JSON.parse(centralResponseBody);
                     if (credentials && !credentials.error) {
                         Logger.log(`[${functionName}]       ✅ Credentials received successfully via refresh.`);
                         credentialsFetched = true;
                     } else if (credentials && credentials.error) {
                         // Definite error from refresh attempt (e.g., record not found, secret mismatch)
                         throw new Error(`Centralized refresh failed: ${credentials.error}`);
                     } else {
                          throw new Error(`Centralized returned unexpected success payload on refresh: ${centralResponseBody}`);
                     }
                 } else {
                      throw new Error(`Failed to refresh credentials. Centralized returned HTTP ${centralStatusCode}. Body: ${centralResponseBody}`);
                 }

            } catch (e) {
                 Logger.log(`[${functionName}]       ❌ Error during 'refreshCredentials' attempt: ${e.message}`);
                 // This is likely the final failure point
                 throw new Error(`Failed to retrieve credentials via cache or refresh: ${e.message}`);
            }
        }

        // --- Validation after potentially successful fetch ---
        if (!credentialsFetched || !credentials) {
             throw new Error("Could not retrieve valid credentials after all attempts.");
        }
        // Validate required credential keys (kept from original code)
        const requiredKeys = ['privateKey', 'apiKey', 'userEmail', 'gcpProjectId', 'serviceAccountEmail'];
        const missingKeys = requiredKeys.filter(key => !(key in credentials));
        if (missingKeys.length > 0) { throw new Error(`Incomplete credentials received. Missing keys: ${missingKeys.join(', ')}`); }
        Logger.log(`[${functionName}]       ✅ Credentials validated successfully.`);
        // --- End Step 2 ---

        // --- Step 3: Store Credentials ---
        // Inside runInitialSetup, in Step 3: Store Credentials
        Logger.log(`[${functionName}] 3. Storing credentials securely...`);
        scriptProps.setProperty('privateKey', credentials.privateKey);
        scriptProps.setProperty('apiKey', credentials.apiKey);                     // Use 'apiKey'
        scriptProps.setProperty('staticMapsApiKey', credentials.staticMapsApiKey); // Use 'staticMapsApiKey'
        scriptProps.setProperty('userEmail', credentials.userEmail);
        scriptProps.setProperty('gcpProjectId', credentials.gcpProjectId);
        scriptProps.setProperty('serviceAccountEmail', credentials.serviceAccountEmail);
        Logger.log(`[${functionName}] Stored properties: apiKey, staticMapsApiKey, userEmail, gcpProjectId, serviceAccountEmail (privateKey hidden)`);



        // --- End Step 3 ---

        // --- Step 4: GCP Binding REMOVED ---
        Logger.log("[${functionName}]    4. Skipping automated GCP project binding (Manual association via editor required).");
        // --- End Step 4 ---

        // --- Step 5: Delete Setup Time Trigger ---
        Logger.log("[${functionName}]    5. Deleting setup time trigger...");
        const deleted = deleteOwnSetupTrigger(); // Deletes the time-based trigger for runInitialSetup
        if (deleted) Logger.log(`[${functionName}]       ✅ Setup time trigger deletion successful.`);
        else Logger.log(`[${functionName}]       ⚠️ Setup time trigger deletion failed or trigger not found.`);
        // --- End Step 5 ---

        // --- Step 6: Mark Setup as Complete (Tentative) ---
        // We mark as true now, but might change if trigger creation fails
        Logger.log("[${functionName}]    6. Marking setup as complete (pending trigger creation).");
        scriptProps.setProperty('SETUP_COMPLETE', 'true');
        overallSuccess = true; // Assume success for now
        // --- End Step 6 ---

        // --- Step 6.5: Create Installable onEdit Trigger ---
        Logger.log(`[${functionName}]    6.5. Creating installable onEdit trigger to call 'handleSheetEdit'...`);
        try {
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            let triggerExists = false;
            const triggers = ScriptApp.getUserTriggers(ss); // Check triggers for *this* spreadsheet
            for (const trigger of triggers) {
                if (trigger.getEventType() === ScriptApp.EventType.ON_EDIT && trigger.getHandlerFunction() === 'handleSheetEdit') {
                    triggerExists = true;
                    Logger.log(`[${functionName}]       Installable trigger for 'handleSheetEdit' already exists.`);
                    break;
                }
            }

            if (!triggerExists) {
                ScriptApp.newTrigger('handleSheetEdit') // Point to the renamed handler function
                    .forSpreadsheet(ss) // Trigger for THIS spreadsheet
                    .onEdit()           // Trigger type is onEdit
                    .create();          // Create the trigger
                Logger.log(`[${functionName}]       ✅ Successfully created installable onEdit trigger for handleSheetEdit.`);
            }
        } catch (err) {
            overallSuccess = false; // Setup wasn't fully successful if trigger fails
            Logger.log(`[${functionName}]       ❌ FAILED to create installable onEdit trigger: ${err.message} - Stack: ${err.stack || 'N/A'}`);
            scriptProps.setProperty('SETUP_COMPLETE', 'trigger_error'); // Update status to reflect trigger issue
            // Decide if you want to re-throw the error to stop execution or just log it
            // throw new Error(`Failed to create necessary onEdit trigger: ${err.message}`); // Option to halt
        }
        // --- End Step 6.5 ---

        // Final Success Logging (Conditional)
        if (overallSuccess) {
            Logger.log('[${functionName}] ✅✅✅ Initial Setup Completed Successfully! ✅✅✅');
        } else {
             Logger.log('[${functionName}] ⚠️⚠️⚠️ Initial Setup Completed with errors (failed to create onEdit trigger). ⚠️⚠️⚠️');
        }

        // --- Step 7: Run Initial Analysis ---
        // Run analysis even if trigger creation failed, as main functionality might still work
        Logger.log(`[${functionName}]    7. Running initial analysis...`);
        // Context initialization should happen within main() or just before
        main(context); // Assuming main() calls initializeContext() internally now
        Logger.log(`[${functionName}]       ✅ Initial analysis complete.`);
        // --- End Step 7 ---

    } catch (error) {
        // Catch errors from steps 1-7 (excluding trigger creation error if not re-thrown)
        Logger.log(`[${functionName}] ❌❌❌ FATAL SETUP ERROR in runInitialSetup: ${error.message}`);
        Logger.log(`   Stack Trace (if available): ${error.stack || 'N/A'}`);
        // Mark setup as failed if a fatal error occurred before completion/trigger step
        if (!scriptProps.getProperty('SETUP_COMPLETE') || scriptProps.getProperty('SETUP_COMPLETE') === 'false') {
             scriptProps.setProperty('SETUP_COMPLETE', 'fatal_error');
        }
    } finally {
        // Optional: Double-check time trigger deletion if an error occurred early
        // deleteOwnSetupTrigger();
        Logger.log(`[${functionName}] Setup process finished.`);
    }
}

/**
 * Checks if a time-driven trigger for runInitialSetup already exists.
 * @return {boolean} True if the trigger exists, false otherwise.
 */
function doesSetupTriggerExist() {
    try {
        const triggers = ScriptApp.getProjectTriggers(); // Needs script.scriptapp scope
        for (const trigger of triggers) {
            if (trigger.getHandlerFunction() === 'runInitialSetup' &&
                trigger.getEventType() === ScriptApp.EventType.CLOCK) {
                Logger.log(`   🔍 Found existing setup trigger ID: ${trigger.getUniqueId()}`);
                return true;
            }
        }
    } catch (err) {
        Logger.log(`   ⚠️ Error checking for existing triggers: ${err.message}. Assuming no trigger exists.`);
    }
    Logger.log("   🔍 No existing setup trigger found.");
    return false;
}

/**
 * Creates the time-driven trigger for runInitialSetup.
 */
function createSetupTrigger() {
    try {
        // Delete potential duplicates first (belt and suspenders)
        deleteOwnSetupTrigger();

        // Create trigger to run in ~60 seconds (more reliable than 30s)
        const trigger = ScriptApp.newTrigger('runInitialSetup') // Needs script.scriptapp scope
            .timeBased()
            .after(60 * 1000)
            .create();
        Logger.log(`   ✅ Successfully created setup trigger ID: ${trigger.getUniqueId()} to run in approx 30 seconds.`);
    } catch (err) {
        Logger.log(`   ❌ Failed to create setup trigger: ${err.message}`);
        throw new Error(`Failed to create the necessary setup trigger: ${err.message}`);
    }
}


/**
 * Finds and deletes the time-driven trigger responsible for calling runInitialSetup.
 * (No changes needed to this function)
 */
function deleteOwnSetupTrigger() {
    let deleted = false;
    try {
        const triggers = ScriptApp.getProjectTriggers(); // Needs script.scriptapp scope
        for (const trigger of triggers) {
            if (trigger.getHandlerFunction() === 'runInitialSetup' && trigger.getEventType() === ScriptApp.EventType.CLOCK) {
                const triggerId = trigger.getUniqueId();
                ScriptApp.deleteTrigger(trigger);
                Logger.log(`   🗑️ Self-deleted setup trigger ID: ${triggerId}`);
                deleted = true;
                break; // Assume only one
            }
        }
        if (!deleted) {
            Logger.log("   ⚠️ No setup trigger found to delete (might have been deleted already or failed to create).");
        }
    } catch (err) {
        Logger.log(`   ❌ Error deleting setup trigger: ${err.message}`);
    }
    return deleted;
}

// === OAuth2 Service Configuration ===
/**
 * Creates and configures the OAuth2 service instance using credentials
 * stored in Script Properties by runInitialSetup.
 * Ensure this uses the correct userSymbol ('OAuth2') and requests necessary scopes.
 */
function getOAuthService() {
    const scriptProps = PropertiesService.getScriptProperties();
    const privateKey = scriptProps.getProperty('privateKey');
    const serviceAccountEmail = scriptProps.getProperty('serviceAccountEmail');
    const impersonatedUserEmail = scriptProps.getProperty('userEmail');

    if (!privateKey || !serviceAccountEmail || !impersonatedUserEmail) {
        Logger.log("[Local OAuth] Credentials (privateKey, serviceAccountEmail, userEmail) not yet available in properties.");
        return null;
    }

    try {
        // !!! Ensure userSymbol 'OAuth2' matches the one in local's appsscript.json !!!
        return OAuth2.createService('googleapis_local') // Use correct library symbol
            .setTokenUrl('https://oauth2.googleapis.com/token')
            .setPrivateKey(privateKey.replace(/\\n/g, '\n'))
            .setIssuer(serviceAccountEmail)
            .setSubject(impersonatedUserEmail)
            .setPropertyStore(scriptProps)
            .setCache(CacheService.getUserCache())
            .setLock(LockService.getUserLock())
            .setScope([ // Scopes needed BY LOCAL SCRIPT for its operations
                'https://www.googleapis.com/auth/spreadsheets',         // For sheet ops
                'https://www.googleapis.com/auth/drive',                // For getFileById().getDescription()
                'https://www.googleapis.com/auth/script.external_request', // For UrlFetchApp (Centralized, Geocoding, Apps Script API)
                'https://www.googleapis.com/auth/script.scriptapp',      // For trigger management (get/delete/create)
                'https://www.googleapis.com/auth/script.projects'        // For Apps Script API (bindSelfToGcpProject)
                // Add 'https://www.googleapis.com/auth/script.container.ui' if UI elements are used
            ].join(' '));
    } catch (err) {
        Logger.log(`[Local OAuth] ❌ Error creating OAuth service: ${err.message} - Check library symbol ('OAuth2'?), credentials, scopes.`);
        return null;
    }
}



/**
 * Normalizes an address string for consistent comparison.
 * Removes punctuation, standardizes abbreviations, trims, and lowers case.
 * @param {string} address
 * @returns {string} normalizedAddress
 */
function normalizeAddress(address) {
  if (!address) return "";

  return String(address)
    .toLowerCase()
    // Remove punctuation except #
    .replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g, "")
    // Standardize common abbreviations
    .replace(/\b(street|str)\b/g, 'st')
    .replace(/\b(road|rd)\b/g, 'rd')
    .replace(/\b(avenue|ave)\b/g, 'ave')
    .replace(/\b(boulevard|blvd)\b/g, 'blvd')
    .replace(/\b(lane|ln)\b/g, 'ln')
    .replace(/\b(place|plz|pl)\b/g, 'pl')
    .replace(/\b(court|ct)\b/g, 'ct')
    .replace(/\b(cove|cv)\b/g, 'cv')
    .replace(/\b(trail|trl)\b/g, 'trl')
    .replace(/\b(drive|dr)\b/g, 'dr')
    .replace(/\b(circle|cir)\b/g, 'cir')
    // Handle unit indicators
    .replace(/\b(unit|apt|suite)\b/g, '#')
    // Handle directionals (optional, uncomment as needed)
    .replace(/\bnorth\b/g, 'n')
    .replace(/\bsouth\b/g, 's')
    .replace(/\beast\b/g, 'e')
    .replace(/\bwest\b/g, 'w')
    // Collapse multiple spaces and trim
    .trim()
    .replace(/\s+/g, ' ');
}

/**
 * Finds the row number in the Preliminary sheet that matches the given simple address.
 * @param {Sheet} preliminarySheet The Google Sheet object for the Preliminary sheet.
 * @param {string} addressToFind The simple property address to search for.
 * @return {number|null} The row number (1-based) if found, otherwise null.
 */
function findRowByAddressInPreliminary(preliminarySheet, addressToFind) {
  const functionName = "findRowByAddressInPreliminary"; // For logging
  const normalizedAddressToFind = normalizeAddress(addressToFind);
  if (!normalizedAddressToFind) {
    Logger.log(`[${functionName}] Error: Cannot search for an empty or invalid address.`);
    return null;
  }
  Logger.log(`[${functionName}] Searching for normalized address: "${normalizedAddressToFind}"`);

  try {
    // Start searching from row 4 (index 3), Column B (index 2)
    const addressColumnRange = preliminarySheet.getRange('B4:B' + preliminarySheet.getLastRow());
    const addressColumnValues = addressColumnRange.getValues();

    for (let i = 0; i < addressColumnValues.length; i++) {
      const currentAddress = addressColumnValues[i][0];
      if (currentAddress) { // Check if the cell is not empty
        const normalizedCurrentAddress = normalizeAddress(currentAddress);
        if (normalizedCurrentAddress === normalizedAddressToFind) {
          const rowNumber = i + 4; // +4 because our range started at row 4
          Logger.log(`[${functionName}] Found match at row ${rowNumber}`);
          return rowNumber;
        }
      }
      // Optimization Removed: Don't break on empty rows prematurely, could miss data lower down.
    }

    Logger.log(`[${functionName}] Address "${addressToFind}" not found in Preliminary sheet column B.`);
    return null; // Address not found

  } catch (error) {
    Logger.log(`[${functionName}] Error searching Preliminary sheet: ${error.message} - Stack: ${error.stack || 'N/A'}`);
    return null;
  }
}

/**
 * Reads the final analysis results from this Local sheet and writes them
 * back to the corresponding row in the Preliminary sheet.
 */
function updatePreliminarySheet() {
  const functionName = "updatePreliminarySheet"; // For logging
  Logger.log(`[${functionName}] Starting update process...`);
  let ss = null; // Initialize spreadsheet variable

  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    const detailedAnalysisSheet = ss.getSheetByName('Detailed Analysis');
    const executiveSummarySheet = ss.getSheetByName('Executive Summary');

    if (!detailedAnalysisSheet || !executiveSummarySheet) {
      throw new Error("Required sheets ('Detailed Analysis' or 'Executive Summary') not found.");
    }

    // 1. Read the Simple Address stored locally
    const simpleAddress = detailedAnalysisSheet.getRange('B6').getValue();
    if (!simpleAddress) {
      throw new Error("Simple address not found in 'Detailed Analysis'!B6.");
    }
    Logger.log(`[${functionName}] Retrieved simple address: "${simpleAddress}"`);

    // 2. Read the calculated metrics from the cells
    // Add error checking/default values if cells might be blank or contain errors
    const netProfit = executiveSummarySheet.getRange('K109').getValue();         
    const simpleROI = executiveSummarySheet.getRange('K111').getValue();         
    const netProfitMargin = executiveSummarySheet.getRange('K110').getValue(); 
    Logger.log(`[${functionName}] Retrieved Metrics from K109, K111, K110 - Net Profit: ${netProfit}, ROI: ${simpleROI}, Margin: ${netProfitMargin}`);

    // Basic validation - ensure we have something to write
     if (netProfit === '' && simpleROI === '' && netProfitMargin === '') {
        Logger.log(`[${functionName}] All metrics are blank. Skipping update to Preliminary.`);
        return; // Don't proceed if all values are blank
     }


    // 3. Open Preliminary Sheet
    Logger.log(`[${functionName}] Opening Preliminary sheet (ID: ${PRELIMINARY_SHEET_ID})...`);
    const preliminarySpreadsheet = SpreadsheetApp.openById(PRELIMINARY_SHEET_ID);
    const preliminarySheet = preliminarySpreadsheet.getSheetByName(PRELIMINARY_SHEET_NAME);
    if (!preliminarySheet) {
      throw new Error(`Sheet "${PRELIMINARY_SHEET_NAME}" not found in Preliminary spreadsheet.`);
    }

    // 4. Find the Target Row using the Simple Address
    const targetRow = findRowByAddressInPreliminary(preliminarySheet, simpleAddress);

    // 5. Write data if row was found
    if (targetRow) {
      Logger.log(`[${functionName}] Writing metrics to Preliminary sheet row ${targetRow}, columns W, X, Y.`);
      // Target columns: W=23, X=24, Y=25
      preliminarySheet.getRange(targetRow, 23, 1, 3).setValues([[
        netProfit,        // Column W
        simpleROI,        // Column X
        netProfitMargin   // Column Y
      ]]);
      SpreadsheetApp.flush(); // Ensure changes are saved
      Logger.log(`[${functionName}] Successfully updated Preliminary sheet row ${targetRow}.`);
    } else {
      // Log error if address wasn't found - this indicates a potential problem
      Logger.log(`[${functionName}] ❌ Error: Could not find row for address "${simpleAddress}" in Preliminary sheet. Update failed.`);
      // Consider more robust error handling here - maybe write to a central error log?
    }

  } catch (error) {
    Logger.log(`[${functionName}] ❌ FATAL ERROR during update: ${error.message} - Stack: ${error.stack || 'N/A'}`);
    // Attempt to log error back to the local sheet if possible
    try {
      if (ss) {
        ss.getSheetByName('Executive Summary').getRange('A150').setValue(`Error updating Preliminary: ${error.message}`); // Example error cell
      }
    } catch (logErr) {
      Logger.log(`[${functionName}] Also failed to write error to local sheet: ${logErr.message}`);
    }
  }
}

/**
 * Handles edits made by users within the configured sheet.
 * Calls main() for a full refresh if Radius (P4) changes.
 * Calls refilterAndAnalyze() for changes in other filter cells (P5, P7, P9, P11).
 * CORRECTED CONTEXT HANDLING & VALIDATION
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e - The edit event object from onEdit trigger.
 */
function handleSheetEdit(e) {
    const functionName = "handleSheetEdit";
    const scriptProps = PropertiesService.getScriptProperties();
    const setupComplete = scriptProps.getProperty('SETUP_COMPLETE');

    // --- Setup validation ---
    if (setupComplete === 'trigger_error') {
        Logger.log(`[${functionName}] Warning: Running after trigger setup error.`);
    } else if (setupComplete !== 'true') {
        Logger.log(`[${functionName}] Setup incomplete. Halting.`);
        return;
    }

    // --- Edit Info ---
    const editedRange = e.range;
    const sheet = editedRange.getSheet();
    const editedCellA1 = editedRange.getA1Notation();
    const sheetName = sheet.getName();
    const editValue = e.value;

    // Use global context for configuration values before full initialization
    const masterSheetName = context.config.MASTER_SHEET_NAME;
    const radiusCell = context.config.COMP_RADIUS_CELL;
    const filterCells = [
        context.config.DATE_FILTER_CELL,
        context.config.AGE_FILTER_CELL,
        context.config.SIZE_FILTER_CELL,
        context.config.SD_MULTIPLIER_CELL
    ];

    // --- Check if the edit is on the correct sheet ---
    if (sheetName !== masterSheetName) {
        return;
    }

    // --- Determine Action based on Edited Cell ---
    if (editedCellA1 === radiusCell) {
        // --- Full Refresh Required ---
        Logger.log(`[${functionName}] Radius change detected: ${sheetName}!${editedCellA1} = ${editValue}. Running FULL refresh (main)...`);
        try {
            initializeContext(); // Update the global context
            // Check if critical parts of the GLOBAL context are now valid
            if (!context || !context.sheet || !context.spreadsheet || !context.config) { // Added !context.config check
                 throw new Error("Global context update failed after initializeContext().");
            }
            main(context); // Pass the updated global context
        } catch (err) {
            Logger.log(`[${functionName}] Error during main() execution: ${err.message}`);
            try { SpreadsheetApp.getUi().alert(`Error during analysis: ${err.message}`); } catch (uiErr) {}
        }

    } else if (filterCells.includes(editedCellA1)) {
        // --- Re-filter and Analysis Only ---
        Logger.log(`[${functionName}] Filter change detected: ${sheetName}!${editedCellA1} = ${editValue}. Running RE-FILTER and ANALYSIS only...`);
         try {
            initializeContext(); // Update the global context
            // Check if critical parts of the GLOBAL context are now valid
             if (!context || !context.sheet || !context.spreadsheet || !context.config) { // Added !context.config check
                 throw new Error("Global context update failed after initializeContext().");
             }
            // Pass the validated GLOBAL context to refilterAndAnalyze
            refilterAndAnalyze(context);
        } catch (err) {
            // Log the specific error from refilterAndAnalyze OR the context check above
            Logger.log(`[${functionName}] Error during refilterAndAnalyze() execution: ${err.message}`);
            try { SpreadsheetApp.getUi().alert(`Error during re-analysis: ${err.message}`); } catch (uiErr) {}
        }

    } else {
        // Logger.log(`[${functionName}] Edit in ${sheetName}!${editedCellA1} not monitored.`);
    }

    Logger.log(`[${functionName}] Finished processing edit for ${editedCellA1}.`);
}

/**
 * Main execution function for the Local sheet analysis.
 * Fetches address, radius, geocodes, finds comps, imports data (incl. Lat/Lng),
 * applies filters (hides rows), and then calls updateAnalysisOutputs.
 * Triggered by runInitialSetup or by handleSheetEdit on Radius change.
 *
 * @param {object} context - The global context object.
 */
function main(context) {
  const functionName = "main";
  Logger.log(`[${functionName}] Starting Main function execution...`);

  try {
    // Initialize and validate context
    initializeContext(); // Assumes global `context` is updated internally
    if (!context || !context.sheet || !context.config || !context.dataSheet || !context.spreadsheet) {
      Logger.log(`[${functionName}] Context initialization failed or missing required components. Exiting.`);
      return;
    }
  } catch (initErr) {
    Logger.log(`[${functionName}] Error during context initialization: ${initErr.message}. Exiting.`);
    return;
  }

  const { sheet, config, spreadsheet } = context;
  Logger.log(` Operating on sheet: "${config.MASTER_SHEET_NAME}"`);

  // Step 1: Retrieve address & radius input
  const address = sheet.getRange(config.ADDRESS_CELL).getValue();
  if (!address) {
    Logger.log(` No address found in ${config.ADDRESS_CELL}. Exiting.`);
    return;
  }
  Logger.log(` Address for geocoding: ${address}`);

  const radius = sheet.getRange(config.COMP_RADIUS_CELL).getValue();
  if (isNaN(radius) || radius <= 0) {
    Logger.log(` Invalid radius in ${config.COMP_RADIUS_CELL}. Exiting.`);
    return;
  }
  Logger.log(` Search radius: ${radius} miles`);

  // Step 2: Geocode the subject property
  const coordinates = getCoordinatesFromAddress(address);
  if (!coordinates) {
    Logger.log(" Failed to geocode subject address. Exiting.");
    return;
  }
  Logger.log(` Subject Coordinates: Lat ${coordinates.lat}, Lng ${coordinates.lng}`);

  // Step 3: Search for comparable properties
  const comps = searchComps(coordinates, radius);

  // Step 4: Import comps into sheet (even if empty, for consistency)
  // This now includes writing data A:P AND setting formulas N:W
  importCompData(comps, context);

  if (Array.isArray(comps) && comps.length > 0) {
    Logger.log(` Imported ${comps.length} comps. Filtering and analyzing...`);

    // Step 5: Apply filtering ONLY
    // clearExistingDataAndFormulas(context);       // KEEP COMMENTED OUT
    applyAllFilters(context);                    // Runs filters (which hide rows)
    // *** Step 5.5: Clear chart data from the rows that were just hidden ***
    clearChartDataForHiddenRows(context);

    // Step 6: Generate outputs: charts, tables, maps, preliminary summary
    updateAnalysisOutputs(context);

  } else {
    // No comps found or imported
    Logger.log(" No comparable properties found/imported. Clearing output areas.");

    try {
      // clearExistingDataAndFormulas(context); // KEEP COMMENTED OUT

      const targetSheet = spreadsheet.getSheetByName('Executive Summary');
      if (targetSheet) {
        targetSheet.getRange('G21:K34').clearContent(); // Map image
        targetSheet.getRange('B23:F32').clearContent(); // Summary table
      }
      // Optional cleanup
      // updateChartWithMargins(context);
      // updatePreliminarySheet();
    } catch (cleanupErr) {
      Logger.log(` Error clearing output areas: ${cleanupErr.message}`);
    }
  }

  Logger.log(`[${functionName}] Main function execution complete.`);
}

/**
 * Applies Standard Deviation filter by hiding rows on the sheet.
 * Reads Price/SqFt directly from visible rows before applying filter.
 *
 * @param {object} context - The global context object.
 * @param {number} sdMultiplier - Number of standard deviations to use for filtering.
 */
function applySDFilter(context, sdMultiplier) {
  const functionName = "applySDFilter";
  const { sheet, config } = context;

  if (!sheet) {
    Logger.log(`[${functionName}] Sheet missing. Skipping SD filter.`);
    return;
  }

  const startRow = config.COMP_RESULTS_START_ROW;
  const endRow = sheet.getLastRow();
  const priceSqFtCol = 14; // Column N (Price/SqFt)
  const pricesPerSqft = [];

  Logger.log(`[${functionName}] Collecting $/SqFt values from visible rows...`);

  if (endRow < startRow || sheet.getMaxColumns() < priceSqFtCol) {
    Logger.log(`[${functionName}] No data rows or Price/SqFt column missing.`);
    return;
  }

  const rowCount = endRow - startRow + 1;
  const values = sheet.getRange(startRow, priceSqFtCol, rowCount, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = startRow + i;
    if (!sheet.isRowHiddenByUser(row)) {
      const val = values[i][0];
      if (val !== null && val !== '' && !isNaN(val) && Number(val) > 0) {
        pricesPerSqft.push(Number(val));
      }
    }
  }

  if (pricesPerSqft.length < 2) {
    Logger.log(`[${functionName}] Not enough visible $/SqFt values (${pricesPerSqft.length}) for SD filtering.`);
    return;
  }

  // --- Calculate Mean and Standard Deviation ---
  const mean = pricesPerSqft.reduce((sum, v) => sum + v, 0) / pricesPerSqft.length;
  const variance = pricesPerSqft.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / pricesPerSqft.length;
  const sd = Math.sqrt(variance);

  const lower = mean - (sdMultiplier * sd);
  const upper = mean + (sdMultiplier * sd);

  Logger.log(` Applying SD filter with bounds: ${lower.toFixed(2)} - ${upper.toFixed(2)} (${pricesPerSqft.length} values)`);

  // --- Apply Filter: Hide out-of-bound rows ---
  const recheckValues = sheet.getRange(startRow, priceSqFtCol, rowCount, 1).getValues();

  for (let i = 0; i < recheckValues.length; i++) {
    const row = startRow + i;
    if (!sheet.isRowHiddenByUser(row)) {
      const val = recheckValues[i][0];
      if (isNaN(val) || val < lower || val > upper) {
        // Logger.log(` Hiding row ${row} with $/SqFt: ${val}`);
        sheet.hideRows(row);
      }
    }
  }

  Logger.log(`[${functionName}] SD filtering completed.`);
}

function clearExistingDataAndFormulas(context) {
    const { sheet, config } = context;
    Logger.log("Clearing existing data and formulas from columns N-W starting at row 33...");
    const lastRow = sheet.getLastRow();
    sheet.getRange(`N${config.COMP_RESULTS_START_ROW}:W${lastRow}`).clearContent();
}

// Refactored applyAllFilters Function (Filtering Actions Enabled)
function applyAllFilters(context) {
    Logger.log("Applying all filters in a nested, sequential order...");

    // Access sheet and config from context
    const { sheet, config } = context;

    if (!sheet) {
        Logger.log(`Sheet "${config.MASTER_SHEET_NAME}" not found.`);
        return;
    }

    const lastRow = sheet.getLastRow();
    const startRow = config.COMP_RESULTS_START_ROW;
    const rowCount = lastRow - startRow + 1;

    // Step 1: Reset visibility for all rows
    if (rowCount > 0) {
         Logger.log(`Showing rows ${startRow} to ${lastRow} before applying filters.`);
         sheet.showRows(startRow, rowCount); // Ensure this is active
         SpreadsheetApp.flush(); // Optional: ensure sheet updates visibility
         Utilities.sleep(200);   // Optional: short pause
    } else {
        Logger.log("No rows to potentially show/filter.");
        return; // Nothing to filter if no rows
    }


    // Step 2: Apply SD Filter First
    const sdMultiplier = sheet.getRange(config.SD_MULTIPLIER_CELL).getValue();
    if (!isNaN(sdMultiplier) && sdMultiplier > 0) {
        Logger.log(`Applying SD filter with multiplier: ${sdMultiplier}`);
        applySDFilter(context, sdMultiplier); // Ensure this is active
    } else {
        Logger.log("Invalid or missing SD multiplier. Skipping SD filtering.");
    }

    // Step 3: Apply each filter sequentially, progressively narrowing down the data set
    Logger.log("Applying Date Filter...");
    applyDateFilter(context); // Ensure this is active

    Logger.log("Applying Age Filter...");
    applyAgeFilter(context);  // Ensure this is active

    Logger.log("Applying Size Filter...");
    applySizeFilter(context); // Ensure this is active

    Logger.log("All filters applied successfully in a nested sequence.");
    SpreadsheetApp.flush(); // Optional: ensure filtering is processed
}


function applyDateFilter(context) {
    const { sheet, config } = context; // Extract sheet and config from the context
    const filterDate = sheet.getRange('P6').getValue();

    if (!(filterDate instanceof Date)) {
        Logger.log("Invalid filter date in P6. Exiting date filter.");
        return;
    }

    Logger.log(`Filtering listings with sale date on or after: ${filterDate}`);
    const lastRow = sheet.getLastRow();
    const saleDatesRange = sheet.getRange(`J${config.COMP_RESULTS_START_ROW}:J${lastRow}`);
    const saleDates = saleDatesRange.getValues();

    for (let i = 0; i < saleDates.length; i++) {
        const saleDate = saleDates[i][0];
        const rowNumber = config.COMP_RESULTS_START_ROW + i;
        if (saleDate instanceof Date && saleDate < filterDate) {
            sheet.hideRows(rowNumber); // Hide rows if the sale date is before the filter date
        }
    }

    Logger.log("Date filter applied successfully.");
}

function applyAgeFilter(context) {
    const { sheet, config } = context; // Extract sheet and config from the context
    const cutoffYear = sheet.getRange('P8').getValue();

    if (isNaN(cutoffYear) || cutoffYear < 1900 || cutoffYear > new Date().getFullYear()) {
        Logger.log("Invalid year input in P8. Exiting age filter.");
        return;
    }

    Logger.log(`Filtering properties built before the year ${cutoffYear}.`);
    const lastRow = sheet.getLastRow();
    const yearBuiltRange = sheet.getRange(`I${config.COMP_RESULTS_START_ROW}:I${lastRow}`);
    const yearBuiltValues = yearBuiltRange.getValues();

    for (let i = 0; i < yearBuiltValues.length; i++) {
        const yearBuilt = yearBuiltValues[i][0];
        const rowNumber = config.COMP_RESULTS_START_ROW + i;
        if (isNaN(yearBuilt) || yearBuilt < cutoffYear) {
            sheet.hideRows(rowNumber); // Hide rows if the property was built before the cutoff year
        }
    }

    Logger.log("Age filter applied successfully.");
}

function applySizeFilter(context) {
    const { sheet, config } = context; // Extract sheet and config from the context
    const subjectSize = sheet.getRange(config.SUBJECT_SIZE_CELL).getValue();

    if (isNaN(subjectSize) || subjectSize <= 0) {
        Logger.log("Invalid subject size in G3. Exiting size filter.");
        return;
    }

    let percentage = sheet.getRange(config.SIZE_FILTER_CELL).getValue();
    if (percentage > 1) {
        percentage = percentage / 100; // Convert to decimal if needed
    }

    if (isNaN(percentage) || percentage < 0) {
        Logger.log("Invalid percentage input in P9. Exiting size filter.");
        return;
    }

    const lowerBound = subjectSize * (1 - percentage);
    const upperBound = subjectSize * (1 + percentage);
    sheet.getRange(config.ANNUNCIATOR_CELL).setValue(`${Math.round(lowerBound)} sqft - ${Math.round(upperBound)} sqft`);
    Logger.log(`Size range: ${Math.round(lowerBound)} sqft to ${Math.round(upperBound)} sqft`);

    const lastRow = sheet.getLastRow();
    const homeSizeRange = sheet.getRange(`G${config.COMP_RESULTS_START_ROW}:G${lastRow}`);
    const homeSizes = homeSizeRange.getValues();

    for (let i = 0; i < homeSizes.length; i++) {
        const homeSize = homeSizes[i][0];
        const rowNumber = config.COMP_RESULTS_START_ROW + i;
        if (isNaN(homeSize) || homeSize < lowerBound || homeSize > upperBound) {
            sheet.hideRows(rowNumber); // Hide rows outside the size range
        }
    }

    Logger.log("Size filter applied successfully.");
}

function calculateTrendlineAndPricePerSqft(context) {
    const { sheet, config } = context; // Extract sheet and config from the context

    // Step 1: Collect comp data for home size and $/sqft from only visible rows
    const compStartRow = config.COMP_RESULTS_START_ROW;
    const lastRow = sheet.getLastRow();
    const dataRange = sheet.getRange(`G${compStartRow}:N${lastRow}`).getValues(); // Corrected to include column N
    const homeSizes = [];
    const pricesPerSqft = [];

    Logger.log("Utilized Home Sizes and $/sqft values for regression:");

    for (let i = 0; i < dataRange.length; i++) {
        const row = compStartRow + i;
        if (sheet.isRowHiddenByUser(row)) continue; // Skip hidden rows

        const homeSize = dataRange[i][0]; // Column G (Home Size)
        const pricePerSqft = dataRange[i][7]; // Column N (Price per SF, corrected index)

        if (isNaN(homeSize) || isNaN(pricePerSqft) || homeSize <= 0 || pricePerSqft <= 0) continue;

        homeSizes.push(homeSize);
        pricesPerSqft.push(pricePerSqft);

        // Log the home size and $/sqft value for verification
        Logger.log(`Home Size: ${homeSize} sqft, $/sqft: ${pricePerSqft}`);
    }

    if (homeSizes.length === 0 || pricesPerSqft.length === 0) {
        Logger.log("No valid comp data found for calculating trendline.");
        return;
    }

    // Step 2: Calculate slope (m) and intercept (b) using linear regression
    const N = homeSizes.length;
    const sumX = homeSizes.reduce((a, b) => a + b, 0);
    const sumY = pricesPerSqft.reduce((a, b) => a + b, 0);
    const sumXY = homeSizes.reduce((sum, x, i) => sum + x * pricesPerSqft[i], 0);
    const sumX2 = homeSizes.reduce((sum, x) => sum + x * x, 0);

    const slope = (N * sumXY - sumX * sumY) / (N * sumX2 - sumX * sumX);
    const intercept = (sumY - slope * sumX) / N;

    Logger.log(`Calculated slope (m): ${slope}`);
    Logger.log(`Calculated intercept (b): ${intercept}`);

    // Step 3: Calculate and set $/sqft for subject properties in rows 3-5
    const subjectRange = sheet.getRange(`G3:G5`).getValues();
    const updatedPrices = [];

    for (let i = 0; i < subjectRange.length; i++) {
        const homeSize = subjectRange[i][0];
        if (isNaN(homeSize) || homeSize <= 0) {
            Logger.log(`Invalid home size in row ${3 + i}. Skipping calculation.`);
            updatedPrices.push([""]); // Push an empty value for invalid entries
            continue;
        }

        // Calculate $/sqft using the trendline equation: y = mx + b
        const pricePerSqft = (slope * homeSize) + intercept;
        updatedPrices.push([pricePerSqft.toFixed(2)]);
        Logger.log(`Calculated $/sqft for home size ${homeSize} sqft: $${pricePerSqft.toFixed(2)}`);
    }

    // Update the prices in one batch
    sheet.getRange(`N3:N5`).setValues(updatedPrices);
}

function updateChartWithMargins(context) {
    const { sheet, config } = context; // Extract sheet and config from the context object
    const charts = sheet.getCharts();

    // Variables to hold the charts we want to modify
    let homePriceChart = null;
    let pricePerSqftChart = null;

    // Loop through charts to find the ones we need
    charts.forEach(chart => {
        const chartTitle = chart.getOptions().get("title");
        if (chartTitle === "Home Price ($) x Home Size (sq.ft.)") {
            homePriceChart = chart;
        } else if (chartTitle === "Home Prices per Square Foot ($) x Home Size (sq.ft.)") {
            pricePerSqftChart = chart;
        }
    });

    if (!homePriceChart || !pricePerSqftChart) {
        Logger.log("One or both charts could not be found. Please check the titles.");
        return;
    }

    // Define start and end of the data range for home sizes, price per sqft, and home prices
    const compStartRow = config.COMP_RESULTS_START_ROW;
    const lastRow = sheet.getLastRow();

    const homeSizes = [];
    const pricesPerSqft = [];
    const homePrices = []; // New array to store home prices

    // Collect data from only visible rows, just like in the regression calculation
    for (let row = compStartRow; row <= lastRow; row++) {
        if (sheet.isRowHiddenByUser(row)) continue; // Skip hidden rows

        const homeSize = sheet.getRange(`G${row}`).getValue();
        const pricePerSqft = sheet.getRange(`N${row}`).getValue();
        const homePrice = sheet.getRange(`K${row}`).getValue(); // Collect home price from column K

        if (isNaN(homeSize) || isNaN(pricePerSqft) || isNaN(homePrice) || homeSize <= 0 || pricePerSqft <= 0 || homePrice <= 0) continue;

        homeSizes.push(homeSize);
        pricesPerSqft.push(pricePerSqft);
        homePrices.push(homePrice); // Add home price to the array
    }

    if (homeSizes.length === 0 || pricesPerSqft.length === 0 || homePrices.length === 0) {
        Logger.log("No valid visible data found for chart resizing.");
        return;
    }

    // Calculate min and max values for visible data
    const minHomeSize = Math.min(...homeSizes);
    const maxHomeSize = Math.max(...homeSizes);
    const minPricePerSqft = Math.min(...pricesPerSqft);
    const maxPricePerSqft = Math.max(...pricesPerSqft);
    const minHomePrice = Math.min(...homePrices);
    const maxHomePrice = Math.max(...homePrices);

    // Apply margins
    const homeSizeMargin = 300;
    const pricePerSqftMargin = 300;
    const homePriceMargin = 300000;

    const adjustedMinHomeSize = Math.max(0, minHomeSize - homeSizeMargin); // Ensure it doesn't go below 0
    const adjustedMaxHomeSize = maxHomeSize + homeSizeMargin;
    const adjustedMinPricePerSqft = Math.max(0, minPricePerSqft - pricePerSqftMargin);
    const adjustedMaxPricePerSqft = maxPricePerSqft + pricePerSqftMargin;
    const adjustedMinHomePrice = Math.max(0, minHomePrice - homePriceMargin);
    const adjustedMaxHomePrice = maxHomePrice + homePriceMargin;

    // Update the first chart for Home Size vs. Home Price
    const updatedHomePriceChart = homePriceChart.modify()
        .setOption("hAxis.minValue", adjustedMinHomeSize)
        .setOption("hAxis.maxValue", adjustedMaxHomeSize)
        .setOption("vAxis.minValue", adjustedMinHomePrice)
        .setOption("vAxis.maxValue", adjustedMaxHomePrice)
        .build();
    sheet.updateChart(updatedHomePriceChart);

    // Update the second chart for Home Size vs. Price per Square Foot
    const updatedPricePerSqftChart = pricePerSqftChart.modify()
        .setOption("hAxis.minValue", adjustedMinHomeSize)
        .setOption("hAxis.maxValue", adjustedMaxHomeSize)
        .setOption("vAxis.minValue", adjustedMinPricePerSqft)
        .setOption("vAxis.maxValue", adjustedMaxPricePerSqft)
        .build();
    sheet.updateChart(updatedPricePerSqftChart);

    Logger.log("Charts updated successfully with margins based on visible data.");
}

// geoComp Functions

/**
 * Fetches geographic coordinates (latitude, longitude) for a given address
 * using the Google Geocoding API and the API key stored in script properties.
 *
 * @param {string} address The street address to geocode.
 * @return {object|null} An object containing {lat: number, lng: number} or null if geocoding fails.
 */
function getCoordinatesFromAddress(address) {
    const functionName = "getCoordinatesFromAddress"; // For logging clarity
    Logger.log(`[${functionName}] Attempting to geocode address: "${address}"`);

    // --- Prerequisite Check: Ensure setup is complete and API key exists ---
    const scriptProps = PropertiesService.getScriptProperties();
    const setupComplete = scriptProps.getProperty('SETUP_COMPLETE');
    const apiKey = scriptProps.getProperty('apiKey'); // Fetched during runInitialSetup

    if (setupComplete !== 'true') {
        Logger.log(`[${functionName}] ❌ Failed: Script setup is not complete. Cannot proceed.`);
        // Optionally alert the user if appropriate context exists
        // SpreadsheetApp.getActiveSpreadsheet().toast('Geocoding failed: Script setup incomplete.');
        return null;
    }

    if (!apiKey) {
        Logger.log(`[${functionName}] ❌ Failed: Geocoding API Key not found in script properties.`);
        // This indicates a setup failure. Alerting might be useful.
        SpreadsheetApp.getActiveSpreadsheet().toast('Geocoding failed: API Key missing. Please check script setup.', 'Configuration Error', 10);
        return null;
    }
    // --- End Prerequisite Check ---

    // --- Construct API Request using API Key ---
    // Using API Key is generally preferred for server-to-server Geocoding calls.
    const baseUrl = "https://maps.googleapis.com/maps/api/geocode/json";
    const encodedAddress = encodeURIComponent(address);
    const url = `${baseUrl}?address=${encodedAddress}&key=${apiKey}`;

    // Log URL without the API key for security
    const logUrl = `${baseUrl}?address=${encodedAddress}&key=***API_KEY***`;
    Logger.log(`[${functionName}] 📡 Requesting Geocoding URL: ${logUrl}`);

    const options = {
        method: "get",
        muteHttpExceptions: true // Allows us to handle API errors gracefully
    };
    // --- End Construct API Request ---

    // --- Execute API Call and Handle Response ---
    let response;
    try {
        response = UrlFetchApp.fetch(url, options);
    } catch (error) {
        // Catch network-level errors (e.g., DNS resolution, timeout)
        Logger.log(`[${functionName}] ❌ Network or UrlFetch Error: ${error.message}`);
        return null;
    }

    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    let data;

    // Try parsing the response
    try {
        data = JSON.parse(responseBody);
    } catch (e) {
        Logger.log(`[${functionName}] ❌ Failed to parse JSON response. Code: ${responseCode}. Body: ${responseBody}`);
        return null;
    }

    // Check response status from the API
    if (responseCode === 200 && data.status === "OK" && data.results && data.results.length > 0) {
        // Success Case
        const location = data.results[0].geometry.location;
        if (location && typeof location.lat === 'number' && typeof location.lng === 'number') {
            Logger.log(`[${functionName}] ✅ Geocoding successful. Coordinates: Lat ${location.lat}, Lng ${location.lng}`);
            return { lat: location.lat, lng: location.lng };
        } else {
            Logger.log(`[${functionName}] ⚠️ Geocoding OK, but location data is invalid in response. Body: ${responseBody}`);
            return null;
        }
    } else {
        // Handle API errors or unexpected statuses
        const apiErrorMessage = data.error_message || `Status: ${data.status}`;
        Logger.log(`[${functionName}] ⚠️ Geocoding API request failed. HTTP Code: ${responseCode}. API Status: ${data.status}. Message: ${apiErrorMessage}`);
        // Specific handling for common issues
        if (data.status === "REQUEST_DENIED") {
             Logger.log(`   Possible causes: Invalid API Key, API not enabled, billing issue.`);
             SpreadsheetApp.getActiveSpreadsheet().toast('Geocoding failed: Request denied by Google Maps. Check API Key/Billing.', 'API Error', 10);
        } else if (data.status === "ZERO_RESULTS") {
             Logger.log(`   The address "${address}" could not be geocoded.`);
             // Decide if this should return null or indicate zero results differently
        }
        // Consider other statuses: OVER_QUERY_LIMIT, INVALID_REQUEST, UNKNOWN_ERROR
        return null;
    }
    // --- End Execute API Call ---
}

function searchComps(coordinates, radius) {
    // Ensure dataSpreadsheet is initialized
    if (!context.dataSpreadsheet) {
        context.dataSpreadsheet = openDataSpreadsheet();
        if (!context.dataSpreadsheet) {
            Logger.log("Data spreadsheet initialization failed.");
            return []; // Exit if the spreadsheet could not be opened
        }
    }

    // Ensure dataSheet is initialized
    if (!context.dataSheet) {
        context.dataSheet = getDataSheet(context.dataSpreadsheet);
        if (!context.dataSheet) {
            Logger.log("Data sheet initialization failed.");
            return []; // Exit if the sheet could not be opened
        }
    }

    const dataValues = context.dataSheet.getDataRange().getValues();
    return findCompsWithinRadius(dataValues, coordinates, radius);
}

// Helper Function: Open the Data Spreadsheet
function openDataSpreadsheet() {
  try {
    return SpreadsheetApp.openById(context.config.DATA_SPREADSHEET_ID);
  } catch (error) {
    Logger.log(`Error opening data spreadsheet: ${error}`);
    return null;
  }
}

function getDataSheet(spreadsheet) {
  const dataSheet = spreadsheet.getSheetByName(context.config.DATA_SHEET_NAME);
  if (!dataSheet) {
    Logger.log(`Sheet "${context.config.DATA_SHEET_NAME}" not found.`);
    return null;
  }
  return dataSheet;
}
// Helper Function: Get Values from the Data Sheet
function getDataSheetValues(sheet) {
    return sheet.getDataRange().getValues();
}

function findCompsWithinRadius(dataValues, coordinates, radius) {
    const comparableProperties = [];
    const functionName = "findCompsWithinRadius";
    Logger.log(`[${functionName}] Searching for comps within ${radius} miles of ${coordinates.lat}, ${coordinates.lng}`);

    for (let i = 1; i < dataValues.length; i++) { // Start from 1 to skip header
        const row = dataValues[i];
        if (!row || row.length <= 41) continue; // Check row length for safety

        const latLngString = row[41]; // Column AP

        if (!latLngString || typeof latLngString !== 'string') continue;

        const [latStr, lngStr] = latLngString.split(',');
        const lat = parseFloat(latStr?.trim());
        const lng = parseFloat(lngStr?.trim());

        if (isNaN(lat) || isNaN(lng)) continue;

        const distance = calculateDistance(coordinates.lat, coordinates.lng, lat, lng);

        if (distance <= radius) {
            comparableProperties.push({
                address:      row[0],  // Col A
                city:         row[2],  // Col C
                state:        row[3],  // Col D
                zip:          row[4],  // Col E
                beds:         row[21], // Col V
                baths:        row[22], // Col W
                buildingSqft: row[23], // Col X
                lotSize:      row[24], // Col Y
                yearBuilt:    row[25], // Col Z
                date:         row[27], // Col AB
                price:        row[36], // Col AK
                distance:     distance,// Calculated
                lat:          lat,     // Parsed Lat
                lng:          lng      // Parsed Lng
                // latLngString is available if needed via property.latLngString
            });
        }
    }
    Logger.log(`[${functionName}] Found ${comparableProperties.length} comparable properties within ${radius} miles.`);
    return comparableProperties;
}

// Helper Function: Calculate Distance Using the Haversine Formula
function calculateDistance(lat1, lng1, lat2, lng2) {
    const toRadians = (degrees) => degrees * (Math.PI / 180);
    const R = 3958.8; // Radius of the Earth in miles

    const dLat = toRadians(lat2 - lat1);
    const dLng = toRadians(lng2 - lng1);

    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
              Math.cos(toRadians(lat1)) * Math.cos(toRadians(lat2)) *
              Math.sin(dLng / 2) * Math.sin(dLng / 2);

    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return R * c; // Distance in miles
}

/**
 * Imports comparable property data into the Sales Comps sheet.
 * Clears previous results, inserts new data (incl. Lat/Lng),
 * MANUALLY sets formulas for N:W, and applies formatting.
 *
 * @param {Array} comparableProperties - Array of comp objects (each must have lat/lng).
 * @param {object} context - The global context object with sheet/config references.
 */
function importCompData(comparableProperties, context) {
    const functionName = "importCompData";

    if (!context || !context.sheet || !context.config) {
        Logger.log(`[${functionName}] Context invalid. Exiting.`);
        return;
    }

    const { sheet, config } = context;
    const numberOfProperties = comparableProperties.length;
    const startDataRow = config.COMP_RESULTS_START_ROW;

    // --- Step 1: Clear old data (A:W) ---
    const clearStartCol = 1; // Column A
    const clearEndCol = 23; // Column W
    const lastRow = sheet.getLastRow();

    if (lastRow >= startDataRow) {
        Logger.log(`[${functionName}] Clearing data A${startDataRow}:${String.fromCharCode(64 + clearEndCol)}${lastRow}`);
        sheet.getRange(startDataRow, clearStartCol, Math.max(1, lastRow - startDataRow + 1), clearEndCol).clearContent();
    } else {
        Logger.log(`[${functionName}] No existing data found to clear.`);
    }

    context.compProperties = [];
    context.visibleAddresses = new Set(); // Assuming you might use this elsewhere

    if (numberOfProperties === 0) {
        Logger.log(`[${functionName}] No comps to import. Skipping insert.`);
        return;
    }

    Logger.log(`[${functionName}] Inserting ${numberOfProperties} comps...`);

    // --- Step 2: Prepare data for insertion (Cols A:P) ---
    const dataToInsert = comparableProperties.map((p, i) => {
        //if (i < 3) { // Keep this debug line commented unless needed
        //    Logger.log(` DEBUG importCompData[${i}]: ${p?.address} | Lat=${p?.lat}, Lng=${p?.lng}`);
        //}
        return [
            p.address || "", // A
            p.city || "", // B
            p.state || "", // C
            p.zip || "", // D
            p.beds || "", // E
            p.baths || "", // F
            p.buildingSqft || "", // G
            p.lotSize || "", // H
            p.yearBuilt || "", // I
            p.date || "", // J
            p.price || "", // K
            "", // L (placeholder)
            p.distance != null ? Number(p.distance).toFixed(2) : "", // M
            "", // N (Price/SqFt formula placeholder - will be set in step 3)
            p.lat != null ? p.lat : "", // O: Latitude
            p.lng != null ? p.lng : "" // P: Longitude
        ];
    });

    try {
        // --- Step 2.5: Write data A:P ---
        const targetRangeAP = sheet.getRange(startDataRow, 1, numberOfProperties, 16); // 16 columns A:P
        Logger.log(`[${functionName}] Writing data to A${startDataRow}:P${startDataRow + numberOfProperties - 1}`);
        targetRangeAP.setValues(dataToInsert);
        SpreadsheetApp.flush(); // Force write to sheet
        Utilities.sleep(500); // Short pause

        // --- Step 2.6: Verification (Optional but keep for now) ---
        Logger.log(`[${functionName}] Verifying written data in O:P...`);
        const latColIndex = 14; // O
        const lngColIndex = 15; // P
        const writtenValues = targetRangeAP.getValues(); // Read back A:P
        let verificationPassed = true;
        let verificationFailures = 0;
        let firstFailureDetails = "";

        for (let k = 0; k < numberOfProperties; k++) {
            const expectedLat = dataToInsert[k][latColIndex];
            const writtenLat = writtenValues[k][latColIndex];
            const expectedLng = dataToInsert[k][lngColIndex];
            const writtenLng = writtenValues[k][lngColIndex];
            const latMatch = (expectedLat === "" && writtenLat === "") || (typeof expectedLat === 'number' && typeof writtenLat === 'number' && Math.abs(expectedLat - writtenLat) < 0.000001) || (String(expectedLat) === String(writtenLat));
            const lngMatch = (expectedLng === "" && writtenLng === "") || (typeof expectedLng === 'number' && typeof writtenLng === 'number' && Math.abs(expectedLng - writtenLng) < 0.000001) || (String(expectedLng) === String(writtenLng));
            if (!latMatch || !lngMatch) {
                verificationPassed = false; verificationFailures++;
                if (!firstFailureDetails) { firstFailureDetails = `Row ${startDataRow + k}: Expected O='${expectedLat}'(type:${typeof expectedLat}), P='${expectedLng}'(type:${typeof expectedLng}) | Got O='${writtenLat}'(type:${typeof writtenLat}), P='${writtenLng}'(type:${typeof writtenLng})`; }
                // break; // Optional: uncomment to stop after first failure
            }
        }
        if (verificationPassed) { Logger.log(`   -> VERIFICATION PASSED for all ${numberOfProperties} rows.`); }
        else { Logger.log(`   -> VERIFICATION FAILED for ${verificationFailures} rows. First failure: ${firstFailureDetails}`); }
        // --- End Verification ---

    } catch (e) {
        Logger.log(`[${functionName}] ERROR inserting comp data (A:P): ${e.message} - Stack: ${e.stack}`);
        throw e;
    }

        // --- Step 3: Copy specific formulas (N and R:W) --- MANUALLY SET FORMULAS ---
    try {
        const formulaSourceRow = 32; // Template row

        // Define the columns that ACTUALLY contain formulas in row 32
        const formulaColsInfo = [
            { col: 14, name: "N" }, // Price/SqFt
            // Skip O (15), P (16), Q (17)
            { col: 18, name: "R" }, // Start of R:W range
            { col: 19, name: "S" },
            { col: 20, name: "T" },
            { col: 21, name: "U" },
            { col: 22, name: "V" },
            { col: 23, name: "W" }  // End of R:W range
        ];

        Logger.log(`[${functionName}] Applying specific formulas for columns: ${formulaColsInfo.map(info => info.name).join(', ')}`);

        // Loop through each COLUMN that needs a formula
        for (const info of formulaColsInfo) {
            const sourceFormulaR1C1 = sheet.getRange(formulaSourceRow, info.col, 1, 1).getFormulaR1C1();

            if (sourceFormulaR1C1) { // Check if the source cell actually has a formula
                Logger.log(`   Applying formula for Col ${info.name} (R1C1: "${sourceFormulaR1C1}") to rows ${startDataRow}-${startDataRow + numberOfProperties - 1}...`);

                // Get the entire destination column range (e.g., N33:N144)
                const targetColumnRange = sheet.getRange(startDataRow, info.col, numberOfProperties, 1);
                try {
                    // Set the formula for the entire column at once using R1C1
                    targetColumnRange.setFormulaR1C1(sourceFormulaR1C1);
                } catch (setFormulaErr) {
                    Logger.log(`   [${functionName}] ERROR setting formula for Col ${info.name}: ${setFormulaErr.message}`);
                }
            } else {
                Logger.log(`   [${functionName}] WARNING: No formula found in template cell ${info.name}${formulaSourceRow}. Skipping column ${info.name}.`);
            }
        }

        Logger.log(`[${functionName}] Finished applying specific formulas.`);
        SpreadsheetApp.flush(); // Ensure formulas are processed

    } catch (formulaErr) {
        Logger.log(`[${functionName}] ERROR setting formulas manually: ${formulaErr.message} - Stack: ${formulaErr.stack}`);
    }
    // --- END Step 3 --- MANUALLY SET SPECIFIC FORMULAS ---


    // --- Step 4: Apply number formatting ---
    try {
        sheet.getRange(startDataRow, 7, numberOfProperties, 2).setNumberFormat('#,##0');         // G:H
        sheet.getRange(startDataRow, 11, numberOfProperties, 1).setNumberFormat('$#,##0');       // K
        sheet.getRange(startDataRow, 13, numberOfProperties, 1).setNumberFormat('0.00');         // M
        sheet.getRange(startDataRow, 15, numberOfProperties, 2).setNumberFormat('0.000000');     // O:P
    } catch (formatErr) {
        Logger.log(`[${functionName}] Warning: Could not apply number formats - ${formatErr.message}`);
    }

    Logger.log(`[${functionName}] Data import completed.`);
}

/**
 * Generates a Google Maps Static API URL with markers.
 * Uses 'visible' parameter for auto zoom/center.
 * Subject: Red 'S'. Comps: Green #'s 1-N.
 * 
 * @param {object} subjectCoords {lat: number, lng: number}.
 * @param {Array<object>} visibleCompCoords Array of {lat: number, lng: number}.
 * @param {string} staticMapApiKey Maps Static API Key.
 * @param {number} width Map width pixels.
 * @param {number} height Map height pixels.
 * @return {string|null} Static Map URL or null.
 */
function generateStaticMapUrl(subjectCoords, visibleCompCoords, staticMapApiKey, width = 640, height = 400) {
  const functionName = "generateStaticMapUrl";

  if (!subjectCoords || !subjectCoords.lat || !subjectCoords.lng || !staticMapApiKey) {
    Logger.log(`[${functionName}] Missing subject coordinates or Static Map API key.`);
    return null;
  }

  const baseUrl = "https://maps.googleapis.com/maps/api/staticmap";
  const params = [];

  // Set map size and type
  params.push(`size=${width}x${height}`);
  params.push(`maptype=roadmap`);

  // --- Construct 'visible' parameter ---
  const visiblePoints = [`${subjectCoords.lat},${subjectCoords.lng}`];
  (visibleCompCoords || []).forEach(comp => {
    if (comp && comp.lat && comp.lng) {
      visiblePoints.push(`${comp.lat},${comp.lng}`);
    }
  });

  if (visiblePoints.length > 1) {
    params.push(`visible=${visiblePoints.join('%7C')}`); // URL-encoded '|'
    Logger.log(`[${functionName}] Using 'visible' parameter with ${visiblePoints.length} points.`);
  } else {
    Logger.log(`[${functionName}] Only subject property coords available. Falling back to center/zoom.`);
    params.push(`center=${subjectCoords.lat},${subjectCoords.lng}`);
    params.push(`zoom=15`);
  }
  // --- End visible param ---

  // --- Markers ---
  // Subject marker (Red S)
  params.push(`markers=color:red%7Clabel:S%7C${subjectCoords.lat},${subjectCoords.lng}`);

  // Comp markers (Green, labeled 1–9, then 'C')
  const maxNumberedMarkers = 9;
  (visibleCompCoords || []).forEach((comp, index) => {
    if (comp && comp.lat && comp.lng) {
      let label = index < maxNumberedMarkers ? (index + 1).toString() : 'C';
      params.push(`markers=color:green%7Clabel:${label}%7C${comp.lat},${comp.lng}`);
    }
  });
  // --- End Markers ---

  // Add API key
  params.push(`key=${staticMapApiKey}`);

  const finalUrl = `${baseUrl}?${params.join('&')}`;
  Logger.log(`[${functionName}] Generated Static Map URL (length ${finalUrl.length}): ${finalUrl}`);

  if (finalUrl.length > 8000) {
    Logger.log(`[${functionName}] WARNING: Generated URL length (${finalUrl.length}) is very long.`);
  }

  return finalUrl;
}

/**
 * Sets the IMAGE formula in the target cell to display the map, 
 * allowing Google Sheets to handle resizing within merged cells.
 * @param {Sheet} targetSheet The sheet object where the image formula should be set.
 * @param {string} mapUrl The URL generated by generateStaticMapUrl.
 * @param {string} targetRangeA1 A1 notation for the top-left cell of the target range (e.g., 'G21').
 * @param {string} mergedRangeA1 Optional: A1 notation for the full merged range (e.g., 'G21:K34') to clear first.
 */
function insertMapIntoSheet(targetSheet, mapUrl, targetRangeA1, mergedRangeA1 = 'G21:K34') {
  const functionName = "insertMapIntoSheet";

  if (!targetSheet || !mapUrl || !targetRangeA1) {
    Logger.log(`[${functionName}] Invalid arguments: targetSheet, mapUrl, or targetRangeA1 missing.`);
    return;
  }

  try {
    // Get the top-left cell where the formula will be placed
    const targetCell = targetSheet.getRange(targetRangeA1);

    // --- Clear Existing Content in Merged Range ---
    if (mergedRangeA1) {
      try {
        Logger.log(`[${functionName}] Clearing content in merged range ${mergedRangeA1}...`);
        targetSheet.getRange(mergedRangeA1).clearContent();
        SpreadsheetApp.flush(); // Ensure content is cleared
        Utilities.sleep(500);   // Short pause to stabilize before inserting
      } catch (clearError) {
        Logger.log(`[${functionName}] Warning: Could not clear range ${mergedRangeA1}. Error: ${clearError.message}`);
      }
    }
    // --- End Clear ---

    // --- Set IMAGE Formula ---
    const imageFormula = `=IMAGE("${mapUrl}", 2)`; // Mode 2 fits merged cell
    Logger.log(`[${functionName}] Setting IMAGE formula in ${targetSheet.getName()}!${targetRangeA1}`);
    // Logger.log(`Formula: ${imageFormula}`); // Uncomment if you want to see full formula

    targetCell.setFormula(imageFormula);
    SpreadsheetApp.flush(); // Helps trigger loading
    // --- End Set Formula ---

    Logger.log(`[${functionName}] IMAGE formula set successfully in ${targetRangeA1}.`);
  } catch (error) {
    Logger.log(`[${functionName}] ERROR setting IMAGE formula in ${targetRangeA1}: ${error.message}`);
    // Optionally: logChangeToCentralLog(...)
  }
}

/**
 * Helper to normalize PRICE values for comparison or calculation.
 * Removes formatting, converts to float, returns NaN if invalid.
 * @param {any} priceValue The price value to normalize.
 * @returns {number|NaN} Normalized float price or NaN.
 */
function normalizePriceForComparison(priceValue) {
  if (priceValue === null || priceValue === undefined) return NaN;

  // Remove $ and , signs
  const cleaned = String(priceValue).replace(/[$,]/g, '').trim();
  if (cleaned === '') return NaN;

  const num = parseFloat(cleaned);
  return isNaN(num) ? NaN : num;
}
/**
 * Populates the comparable properties table (B23:F32) on the Executive Summary sheet
 * by reading data directly from the visible rows on the 'Sales Comps' sheet.
 * Corrected Column Mapping: B=Address, C=Size, D=SalePrice, E=Distance, F=Price/SqFt
 * @param {object} context The global context object.
 */
function populateExecutiveSummaryCompsTable(context) {
    const functionName = "populateExecutiveSummaryCompsTable";
    Logger.log(`[${functionName}] Starting...`);

    try {
        if (!context || !context.spreadsheet || !context.sheet) {
            Logger.log(`[${functionName}] Context or essential sheet objects missing. Skipping.`);
            return;
        }

        const salesCompSheet = context.sheet; // The 'Sales Comps' sheet
        const targetSheetName = 'Executive Summary';
        const targetSheet = context.spreadsheet.getSheetByName(targetSheetName);
        if (!targetSheet) { Logger.log(`[${functionName}] Target sheet "${targetSheetName}" not found. Skipping.`); return; }

        // --- Read data DIRECTLY from VISIBLE rows on Sales Comps sheet ---
        const startDataRow = context.config.COMP_RESULTS_START_ROW;
        const lastDataRow = salesCompSheet.getLastRow();
        const visibleCompsData = []; // To store data read from sheet

        // Define column numbers (1-based)
        const addressCol = 1;     // A
        const sizeCol = 7;        // G
        const priceCol = 11;      // K
        const distanceCol = 13;   // M
        const priceSqFtCol = 14;  // N
        const maxColToRead = priceSqFtCol; // Read up to Column N

        if (lastDataRow >= startDataRow) {
            const numRowsToCheck = lastDataRow - startDataRow + 1;
            Logger.log(`   Checking ${numRowsToCheck} rows on Sales Comps sheet for visibility...`);
            // Read the relevant range (A:N) for all potential rows
            const allData = salesCompSheet.getRange(startDataRow, 1, numRowsToCheck, maxColToRead).getValues();

            for (let i = 0; i < allData.length; i++) {
                const currentRow = startDataRow + i;
                // Check visibility FIRST
                if (!salesCompSheet.isRowHiddenByUser(currentRow)) {
                    const rowData = allData[i];
                    // Extract data using correct indices (0-based)
                    visibleCompsData.push({
                        address:      rowData[addressCol - 1],
                        homeSize:     rowData[sizeCol - 1],
                        salePrice:    rowData[priceCol - 1],
                        distance:     rowData[distanceCol - 1],
                        pricePerSqft: rowData[priceSqFtCol - 1]
                    });
                }
            }
            Logger.log(`   Found ${visibleCompsData.length} visible comps on sheet.`);
        } else {
            Logger.log(`   No data rows found on Sales Comps sheet.`);
        }
        // --- End Reading Data ---

        // Sort visible comps by distance (ascending)
        visibleCompsData.sort((a, b) => (parseFloat(a.distance) || Infinity) - (parseFloat(b.distance) || Infinity));

        // Take top 10 comps
        const compsToDisplay = visibleCompsData.slice(0, 10);
        Logger.log(`   Taking top ${compsToDisplay.length} comps to display.`);

        // Format data for the table B23:F32
        const tableData = compsToDisplay.map((comp) => {
            return [
                comp.address || "N/A",
                comp.homeSize || null,
                comp.salePrice || null,
                comp.distance != null ? Number(comp.distance).toFixed(2) : null,
                comp.pricePerSqft || null
            ];
        });

        // Clear destination range and write data
        const startRow = 23; const startCol = 2; const numCols = 5; // B:F
        const targetRangeA1 = `B${startRow}:F${startRow + 9}`; // B23:F32
        Logger.log(`   Clearing existing data in ${targetSheetName}!${targetRangeA1}`);
        targetSheet.getRange(targetRangeA1).clearContent();
        SpreadsheetApp.flush(); Utilities.sleep(200);

        if (tableData.length > 0) {
            const writeRange = targetSheet.getRange(startRow, startCol, tableData.length, numCols);
            Logger.log(`   Writing ${tableData.length} rows of comp data...`);
            writeRange.setValues(tableData);

            // Apply Number Formatting
            try {
                targetSheet.getRange(startRow, 3, tableData.length, 1).setNumberFormat('#,##0');      // C: Size
                targetSheet.getRange(startRow, 4, tableData.length, 1).setNumberFormat('$#,##0');      // D: Price
                targetSheet.getRange(startRow, 5, tableData.length, 1).setNumberFormat('0.00');       // E: Distance
                targetSheet.getRange(startRow, 6, tableData.length, 1).setNumberFormat('$#,##0.00'); // F: Price/SqFt
            } catch (formatErr) { Logger.log(`   Warning: Could not apply number formats - ${formatErr.message}`); }
            Logger.log(`[${functionName}] Comp table populated successfully.`);
        } else { Logger.log(`[${functionName}] No visible comps to display in table.`); }

    } catch (error) { Logger.log(`[${functionName}] ERROR: ${error.message} - Stack: ${error.stack || 'N/A'}`); }
}

/**
 * Performs final output tasks including calculations, chart updates, map generation,
 * and table population. Pushes results back to the Preliminary sheet.
 * All data is read directly from the sheet based on row visibility.
 * ADDED DETAILED LOGGING FOR MAP COORDINATE COLLECTION
 *
 * @param {object} context - The global context object.
 */
function updateAnalysisOutputs(context) {
    const functionName = "updateAnalysisOutputs";
    Logger.log(`[${functionName}] Starting final calculations and outputs...`);

    try {
        // --- Step 0: Context validation ---
        const isValidContext = context && context.sheet && context.config && context.spreadsheet;
        if (!isValidContext) {
            Logger.log(`[${functionName}] Missing required context components. Skipping.`);
            return;
        }

        // --- Step 1: Trendline & Subject $/SqFt calculation ---
        Logger.log(` Running calculations (trendline, subject $/sqft)...`);
        calculateTrendlineAndPricePerSqft(context); // Uses visibility internally

        // --- Step 2: Chart Update ---
        Logger.log(` Updating charts...`);
        updateChartWithMargins(context); // Uses visibility internally

        // --- Step 3: Executive Summary Comp Table ---
        try {
            Logger.log(` Populating Executive Summary comp table...`);
            populateExecutiveSummaryCompsTable(context); // Uses sheet visibility
        } catch (tableError) {
            Logger.log(` Error in populateExecutiveSummaryCompsTable: ${tableError.message}`);
        }

        // --- Step 4: Static Map Generation ---
        try {
            Logger.log(" Generating static map...");
            const scriptProps = PropertiesService.getScriptProperties();
            const staticMapKey = scriptProps.getProperty('staticMapsApiKey');
            const subjectAddress = context.sheet.getRange(context.config.ADDRESS_CELL).getValue();
            const coordinates = getCoordinatesFromAddress(subjectAddress);

            if (!staticMapKey) {
                Logger.log(" WARNING: Missing 'staticMapsApiKey'. Map skipped.");
            } else if (!coordinates?.lat || !coordinates?.lng) {
                Logger.log(" WARNING: Missing subject coordinates. Map skipped.");
            } else {
                // --- Read visible comp coordinates directly from the sheet ---
                const salesCompSheet = context.sheet;
                const startRow = context.config.COMP_RESULTS_START_ROW;
                const lastRow = salesCompSheet.getLastRow();
                const latCol = 15; // Column O
                const lngCol = 16; // Column P
                const visibleCompCoords = []; // Array to hold {lat, lng} objects

                if (lastRow >= startRow && salesCompSheet.getMaxColumns() >= lngCol) {
                    const rowCount = lastRow - startRow + 1;
                    // ---- START: Added Logging ----
                    Logger.log(`[${functionName}-MapCoords] Reading coords from ${salesCompSheet.getName()}!O${startRow}:P${lastRow}`);
                    const coordsData = salesCompSheet.getRange(startRow, latCol, rowCount, 2).getValues(); // O:P

                    Logger.log(`[${functionName}-MapCoords] Iterating through ${coordsData.length} potential rows to find visible coords...`);
                    // ---- END: Added Logging ----
                    for (let i = 0; i < coordsData.length; i++) {
                        const row = startRow + i;
                        // ---- START: Added Logging ----
                        const isHidden = salesCompSheet.isRowHiddenByUser(row);
                        const lat = coordsData[i][0]; // Raw value from O
                        const lng = coordsData[i][1]; // Raw value from P

                        // Log details for EVERY row checked BEFORE the visibility/validity check
                        Logger.log(`   [${functionName}-MapCoords] Checking row ${row}: Hidden=${isHidden}, Raw Lat='${lat}'(type:${typeof lat}), Raw Lng='${lng}'(type:${typeof lng})`);
                        // ---- END: Added Logging ----

                        if (!isHidden) { // Check visibility first
                            // Validate the coordinates now AFTER checking visibility
                            const isValidCoord = lat != null && lng != null && lat !== '' && lng !== '' && !isNaN(Number(lat)) && !isNaN(Number(lng)); // Check not null, not empty string, and numeric
                            // ---- START: Added Logging ----
                            if (isValidCoord) {
                                // Log successful validation and the values being pushed
                                Logger.log(`      -> Row ${row} is VISIBLE and coords VALID. Pushing {lat: ${Number(lat)}, lng: ${Number(lng)}}`);
                                visibleCompCoords.push({ lat: Number(lat), lng: Number(lng) }); // Use Number() for conversion just in case
                            } else {
                                // Log failure reason
                                Logger.log(`      -> Row ${row} is VISIBLE but coords INVALID (isNull=${lat==null}/${lng==null}, isEmpty=${lat===''}/${lng===''}, isNaN=${isNaN(Number(lat))}/${isNaN(Number(lng))}).`);
                            }
                            // ---- END: Added Logging ----
                        }
                        // If hidden, the loop just continues to the next row
                    }
                    // ---- START: Added Logging ----
                    Logger.log(`[${functionName}-MapCoords] Finished iteration over sheet rows.`);
                    // ---- END: Added Logging ----
                } else {
                    // Log reasons for not reading data
                    if (lastRow < startRow) Logger.log(`[${functionName}-MapCoords] No data rows found (lastRow ${lastRow} < startRow ${startRow}).`);
                    else Logger.log(`[${functionName}-MapCoords] WARNING: Lat/Lng columns (O:P) not found or incomplete (max cols ${salesCompSheet.getMaxColumns()} < ${lngCol}).`);
                }

                // This log reflects the final count of items actually pushed to the array
                Logger.log(` Found ${visibleCompCoords.length} visible comps with valid coordinates added to map array.`);

                // Pass the collected coordinates to the map function
                const mapUrl = generateStaticMapUrl(coordinates, visibleCompCoords, staticMapKey);
                const targetSheet = context.spreadsheet.getSheetByName('Executive Summary');
                const targetRange = 'G21';

                if (mapUrl && targetSheet) {
                    insertMapIntoSheet(targetSheet, mapUrl, targetRange);
                } else {
                    Logger.log(" Skipping map insertion (no map URL or sheet missing / mapUrl is null).");
                    try {
                        if (targetSheet) targetSheet.getRange('G21:K34').clearContent();
                    } catch (e) {
                        Logger.log(" Error clearing map area: " + e.message);
                    }
                }
            }
        } catch (mapError) {
            Logger.log(` ERROR during map processing: ${mapError.message} - Stack: ${mapError.stack || 'N/A'}`);
        }

        // --- Step 5: Update Preliminary Sheet ---
        try {
            Logger.log(" Flushing spreadsheet and updating Preliminary sheet...");
            SpreadsheetApp.flush();
            Utilities.sleep(2000); // Consider if this sleep is necessary
            updatePreliminarySheet();
        } catch (updateErr) {
            Logger.log(` Error updating Preliminary sheet: ${updateErr.message}`);
        }

        Logger.log(`[${functionName}] Final output processing complete.`);

    } catch (error) {
        Logger.log(`[${functionName}] FATAL ERROR: ${error.message} - Stack: ${error.stack || 'N/A'}`);
    }
}

/**
 * Re-applies filters, formulas, clears hidden chart data, and updates outputs.
 * Added delays between steps for stability.
 *
 * @param {object} context - The global context object, assumed to be valid and initialized.
 */
function refilterAndAnalyze(context) {
    const functionName = "refilterAndAnalyze";
    Logger.log(`[${functionName}] Starting re-filter and analysis based on existing data...`);

    try {
        const { sheet, config, spreadsheet } = context;
        if (!sheet || !config || !spreadsheet) {
             throw new Error("Invalid context object received by refilterAndAnalyze.");
        }

        // Step 1: Apply filtering
        Logger.log(`[${functionName}] --- Applying Filters ---`);
        applyAllFilters(context);
        Utilities.sleep(1000); // *** ADDED: Pause 1 second after filtering ***

        // Step 1.5: Re-apply formulas to currently VISIBLE rows
        Logger.log(`[${functionName}] --- Applying Formulas to Visible Rows ---`);
        const startDataRow = config.COMP_RESULTS_START_ROW;
        const lastDataRow = sheet.getLastRow();
        const visibleRows = [];
        if (lastDataRow >= startDataRow) {
             Logger.log(`[${functionName}] Identifying visible rows between ${startDataRow} and ${lastDataRow}...`);
             for (let r = startDataRow; r <= lastDataRow; r++) {
                 if (!sheet.isRowHiddenByUser(r)) {
                     visibleRows.push(r);
                 }
             }
             Logger.log(`[${functionName}] Found ${visibleRows.length} visible rows to apply formulas to.`);
             applyFormulasToRows(sheet, visibleRows, 32); // Apply formulas
             Utilities.sleep(1000); // *** ADDED: Pause 1 second after applying formulas ***
        } else {
             Logger.log(`[${functionName}] No data rows found, skipping formula application.`);
        }

        // Step 2: Clear chart data from the rows that are now hidden
        Logger.log(`[${functionName}] --- Clearing Chart Data from Hidden Rows ---`);
        clearChartDataForHiddenRows(context);
        Utilities.sleep(1000); // *** ADDED: Pause 1 second after clearing hidden rows ***

        // Step 3: Generate outputs using only currently visible data
        Logger.log(`[${functionName}] --- Generating Final Outputs ---`);
        updateAnalysisOutputs(context);

    } catch (err) {
        Logger.log(`[${functionName}] FATAL ERROR during execution: ${err.message} - Stack: ${err.stack || 'N/A'}`);
         throw err;
    }

    Logger.log(`[${functionName}] Re-filter and analysis complete.`);
}

/**
 * Clears specific chart data columns (R, T, W) for rows that are hidden
 * after filtering has been applied. Does NOT clear O:P or other source data.
 * Should be called AFTER applyAllFilters and BEFORE updateAnalysisOutputs.
 * @param {object} context The global context object.
 */
function clearChartDataForHiddenRows(context) {
    const functionName = "clearChartDataForHiddenRows";
    const { sheet, config } = context;
    const startRow = config.COMP_RESULTS_START_ROW; // Typically 33
    const lastRow = sheet.getLastRow();

    if (!sheet) {
        Logger.log(`[${functionName}] Sheet not found in context. Skipping.`);
        return;
    }
     if (lastRow < startRow) {
        Logger.log(`[${functionName}] No data rows (${startRow}+) to check.`);
        return;
    }

    // Define the columns used by the charts that need clearing for hidden rows
    const chartDataCols = [ // Updated columns
        18, // Column R
        20, // Column T
        23  // Column W
    ];

    Logger.log(`[${functionName}] Clearing chart data (Cols R, T, W) for hidden rows from ${startRow} to ${lastRow}...`);
    let hiddenRowsClearedCount = 0;
    let errorCount = 0;

    // Iterate through all potentially relevant rows
    for (let r = startRow; r <= lastRow; r++) {
        try {
            // Check if the current row is hidden
            if (sheet.isRowHiddenByUser(r)) {
                // If hidden, clear the content of the specified chart columns for this row
                chartDataCols.forEach(colNum => {
                    if (colNum <= sheet.getMaxColumns()) {
                         sheet.getRange(r, colNum).clearContent();
                    }
                });
                hiddenRowsClearedCount++;
            }
        } catch (e) {
             errorCount++;
             Logger.log(`   [${functionName}] Error checking/clearing row ${r}: ${e.message}`);
        }
    }

    if (hiddenRowsClearedCount > 0 || errorCount > 0) {
        if (hiddenRowsClearedCount > 0) {
             SpreadsheetApp.flush(); // Apply the clearing actions
             Logger.log(`[${functionName}] Finished clearing chart data for ${hiddenRowsClearedCount} hidden rows.`);
        }
        if (errorCount > 0) {
             Logger.log(`[${functionName}] Encountered ${errorCount} errors during row processing.`);
        }
    } else {
         Logger.log(`[${functionName}] No hidden rows found to clear chart data from.`);
    }
}
/**
 * Applies the template formulas from row 32 (N, R:W) to a specified list of rows.
 * Uses R1C1 notation for automatic reference adjustment.
 * Designed to be called after filtering to populate newly visible rows.
 *
 * @param {Sheet} sheet The sheet object ('Sales Comps').
 * @param {number[]} targetRows An array of row numbers (1-based) to apply formulas to.
 * @param {number} formulaSourceRow The row number containing the template formulas (e.g., 32).
 */
function applyFormulasToRows(sheet, targetRows, formulaSourceRow) {
    const functionName = "applyFormulasToRows";
    if (!targetRows || targetRows.length === 0) {
        Logger.log(`[${functionName}] No target rows provided. Skipping.`);
        return;
    }
     if (!sheet) {
        Logger.log(`[${functionName}] Invalid sheet object provided. Skipping.`);
        return;
    }

    // Define the columns that ACTUALLY contain formulas in the template row
    const formulaColsInfo = [
        { col: 14, name: "N" }, // Price/SqFt
        // Skip O (15), P (16), Q (17) - Assuming these don't have formulas in row 32
        { col: 18, name: "R" }, // Home Size (Usually a direct reference like =G33)
        { col: 19, name: "S" }, // Check if S32 has a formula
        { col: 20, name: "T" }, // Comp Sale Price (Usually a direct reference like =K33)
        { col: 21, name: "U" }, // Check if U32 has a formula
        { col: 22, name: "V" }, // Check if V32 has a formula
        { col: 23, name: "W" }  // Comp Price per SF (Usually calculated)
    ];

    Logger.log(`[${functionName}] Applying formulas for columns: ${formulaColsInfo.map(info => info.name).join(', ')} to ${targetRows.length} rows.`);
    let formulaErrors = 0;

    // Get all source formulas at once
    const sourceFormulasR1C1 = {};
    formulaColsInfo.forEach(info => {
        try {
            const formula = sheet.getRange(formulaSourceRow, info.col).getFormulaR1C1();
            if (formula) {
                 sourceFormulasR1C1[info.col] = formula;
            } else {
                 Logger.log(`   [${functionName}] WARNING: No formula found in template cell ${info.name}${formulaSourceRow}. Will skip applying for Col ${info.name}.`);
                 sourceFormulasR1C1[info.col] = null; // Mark as null if no formula
            }
        } catch (e) {
             Logger.log(`   [${functionName}] ERROR reading formula from ${info.name}${formulaSourceRow}: ${e.message}`);
             sourceFormulasR1C1[info.col] = null; // Mark as null on error
        }
    });


    // Apply formulas row by row ONLY to the specified targetRows
    targetRows.forEach(rowNum => {
        formulaColsInfo.forEach(info => {
            const formulaToApply = sourceFormulasR1C1[info.col];
            if (formulaToApply) { // Only apply if a valid formula was found for this column
                 try {
                    sheet.getRange(rowNum, info.col).setFormulaR1C1(formulaToApply);
                } catch (setErr) {
                    Logger.log(`   [${functionName}] ERROR applying formula to ${info.name}${rowNum}: ${setErr.message}`);
                    formulaErrors++;
                 }
             }
        });
    });

    if (formulaErrors === 0) {
        Logger.log(`[${functionName}] Finished applying formulas successfully.`);
    } else {
        Logger.log(`[${functionName}] Finished applying formulas with ${formulaErrors} errors.`);
    }
    SpreadsheetApp.flush(); // Apply changes
}

/**
 * Creates a new Google Slides presentation by copying a template, placing it
 * in the sheet's parent folder, and populating Slide 4 with data and charts
 * from the active Local Google Sheet.
 */
function createPresentationFromSheet() {
    const functionName = "createPresentationFromSheet";
    const ui = SpreadsheetApp.getUi();
    Logger.log(`[${functionName}] Starting presentation generation...`);

    // --- Input Validation ---
    if (!SLIDES_TEMPLATE_ID || SLIDES_TEMPLATE_ID === 'YOUR_SLIDES_TEMPLATE_ID_GOES_HERE') {
         ui.alert("Slides Template ID is not set in the script constants.");
         Logger.log(`[${functionName}] Error: SLIDES_TEMPLATE_ID constant not set.`);
         return;
    }
     if (!CHART_SHEET_NAME || CHART_SHEET_NAME === 'Name of Sheet With Charts' || !PIE_CHART_1_TITLE || PIE_CHART_1_TITLE === 'Exact Title of First Pie Chart' || !PIE_CHART_2_TITLE || PIE_CHART_2_TITLE === 'Exact Title of Second Pie Chart') {
         ui.alert("Chart information (Sheet Name, Chart Titles) is not set in the script constants.");
         Logger.log(`[${functionName}] Error: Chart constants not set.`);
         return;
    }
    const TARGET_SLIDE_INDEX = 3; // Slide 4 is index 3 (0-based)

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const detailedAnalysisSheet = ss.getSheetByName('Detailed Analysis');
        const chartSheet = ss.getSheetByName(CHART_SHEET_NAME); // Get the sheet containing charts

        if (!detailedAnalysisSheet) { throw new Error("Sheet 'Detailed Analysis' not found."); }
        if (!chartSheet) { throw new Error(`Sheet named '${CHART_SHEET_NAME}' (for charts) not found.`); }

        // --- 1. Read Data from Sheet ---
        Logger.log("   Reading data from Detailed Analysis sheet...");
        const propertyAddressSimple = detailedAnalysisSheet.getRange('B6').getValue(); // Use simple address for naming
        if (!propertyAddressSimple) { throw new Error("Simple address not found in 'Detailed Analysis'!B6."); }

        // Get display values to preserve formatting (e.g., %, $)
        const targetIRR = detailedAnalysisSheet.getRange('B153').getDisplayValue();
        const targetROI = detailedAnalysisSheet.getRange('B151').getDisplayValue();
        const targetMultiple = detailedAnalysisSheet.getRange('B152').getDisplayValue();
        const netProfit = detailedAnalysisSheet.getRange('B138').getDisplayValue();
        const grossProfit = detailedAnalysisSheet.getRange('B135').getDisplayValue();

        Logger.log(`   Data read: IRR=${targetIRR}, ROI=${targetROI}, Multiple=${targetMultiple}, Net=${netProfit}, Gross=${grossProfit}`);

        // --- 2. Get Drive Folder ---
        Logger.log("   Identifying Drive folder...");
        const sheetFile = DriveApp.getFileById(ss.getId());
        const parentFolders = sheetFile.getParents();
        if (!parentFolders.hasNext()) { throw new Error("Cannot determine parent folder of this spreadsheet."); }
        const parentFolder = parentFolders.next();
        Logger.log(`   Found parent folder: ${parentFolder.getName()} (ID: ${parentFolder.getId()})`);

        // --- 3. Copy Template & Rename ---
        Logger.log(`   Copying Slides template (ID: ${SLIDES_TEMPLATE_ID})...`);
        const templateFile = DriveApp.getFileById(SLIDES_TEMPLATE_ID);
        const newFileName = `${propertyAddressSimple} - Investor Summary`; // Naming convention
        const newPresentationFile = templateFile.makeCopy(newFileName, parentFolder);
        const newPresentationId = newPresentationFile.getId();
        const newPresentationUrl = newPresentationFile.getUrl();
        Logger.log(`   Created new presentation: ${newFileName} (ID: ${newPresentationId})`);
        Utilities.sleep(1000); // Brief pause after copy

        // --- 4. Open and Get Target Slide ---
        Logger.log(`   Opening new presentation to populate...`);
        const presentation = SlidesApp.openById(newPresentationId);
        const slides = presentation.getSlides();

        if (slides.length <= TARGET_SLIDE_INDEX) {
            throw new Error(`Template does not have enough slides. Target slide index ${TARGET_SLIDE_INDEX} (Slide ${TARGET_SLIDE_INDEX + 1}) is out of bounds.`);
        }
        const targetSlide = slides[TARGET_SLIDE_INDEX]; // Get Slide 4
        Logger.log(`   Targeting Slide ${TARGET_SLIDE_INDEX + 1} (ID: ${targetSlide.getObjectId()}).`);

        // --- 5. Replace Text Placeholders (Reverted to replaceAllText) ---
        Logger.log(`   Replacing text placeholders on Slide 4 using replaceAllText...`);

        // Use replaceAllText for each placeholder. Bolding depends on template placeholder formatting.
        targetSlide.replaceAllText('{{TARGET_IRR}}', targetIRR || 'N/A');
        targetSlide.replaceAllText('{{TARGET_ROI}}', targetROI || 'N/A');
        targetSlide.replaceAllText('{{TARGET_MULTIPLE}}', targetMultiple || 'N/A');
        targetSlide.replaceAllText('{{NET_PROFIT}}', netProfit || 'N/A');
        targetSlide.replaceAllText('{{GROSS_PROFIT}}', grossProfit || 'N/A');
        // Add more calls if you have other placeholders to replace on this slide

        SpreadsheetApp.flush(); // Allow SlidesApp changes to process
        Utilities.sleep(500);   // Small pause can sometimes help reliability
        // --- END Step 5 --- Reverted ---

        // --- 6. Find and Embed Charts ---
        Logger.log(`   Finding charts on sheet '${CHART_SHEET_NAME}'...`);
        const charts = chartSheet.getCharts();
        let chart1 = null;
        let chart2 = null;

        // Find the charts by their exact titles
        for (const chart of charts) {
            const title = chart.getOptions().get('title');
            if (title === PIE_CHART_1_TITLE) {
                chart1 = chart;
                Logger.log(`      Found Chart 1: "${PIE_CHART_1_TITLE}"`);
            } else if (title === PIE_CHART_2_TITLE) {
                chart2 = chart;
                Logger.log(`      Found Chart 2: "${PIE_CHART_2_TITLE}"`);
            }
            if (chart1 && chart2) break;
        }

        // --- Define Chart Positions and Sizes (ADJUST VALUES BELOW) ---
        const slideWidth = presentation.getPageWidth();
        const slideHeight = presentation.getPageHeight();

        // ** Step A: Define Desired WIDTH **
        const desiredChartWidth = slideWidth * 0.45;  // <<< ADJUST this multiplier (e.g., 0.35, 0.4)

        // ** Step B: Define Desired Aspect Ratio (Height relative to Width) **
        const aspectRatio = .75; // <<< ADJUST this based on chart shape (e.g., 1.0 for square pie, 0.75 standard)
        const calculatedChartHeight = desiredChartWidth * aspectRatio; // Calculate the height based on width and desired ratio

        // ** Step C: Calculate Vertical Position (Bottom Alignment) **
        const bottomMargin = 10; // <<< ADJUST points from bottom edge
        // Use the CALCULATED height for positioning
        const topPosition = slideHeight - calculatedChartHeight - bottomMargin;

        // ** Step D: Calculate Horizontal Positions (Centering with Gap) **
        const gap = -25; // <<< ADJUSTABLE: Points of space BETWEEN the two charts
        const totalChartsWidth = 2 * desiredChartWidth + gap;
        const sideMarginCalculated = (slideWidth - totalChartsWidth) / 2;
        const leftMargin = Math.max(10, sideMarginCalculated); // Ensure minimum 10pt margin

        const leftPosition1 = leftMargin;
        const leftPosition2 = leftMargin + desiredChartWidth + gap;

        // Log the definite width and height being used
        Logger.log(`   Targeting Pos: Chart1(L:${leftPosition1.toFixed(1)}, T:${topPosition.toFixed(1)}), Chart2(L:${leftPosition2.toFixed(1)}, T:${topPosition.toFixed(1)}), Size(W:${desiredChartWidth.toFixed(1)}, H:${calculatedChartHeight.toFixed(1)})`);

        // ** Step E: Embed Charts specifying BOTH Width and calculated Height **

        // Embed Chart 1
        if (chart1) {
            Logger.log(`   Embedding Chart 1 ("${PIE_CHART_1_TITLE}") onto Slide 4...`);
            try {
                // Insert specifying BOTH calculated width and height
                const image1 = targetSlide.insertSheetsChartAsImage(chart1, leftPosition1, topPosition, desiredChartWidth, calculatedChartHeight);
                Logger.log(`      Embedded Chart 1 successfully.`);
            } catch(chartErr) {
                 Logger.log(`      ERROR embedding Chart 1: ${chartErr.message}`);
            }
        } else {
            Logger.log(`   WARNING: Chart 1 "${PIE_CHART_1_TITLE}" not found on sheet "${CHART_SHEET_NAME}".`);
        }

        // Embed Chart 2
        if (chart2) {
             Logger.log(`   Embedding Chart 2 ("${PIE_CHART_2_TITLE}") onto Slide 4...`);
             try {
                // Insert specifying BOTH calculated width and height
                const image2 = targetSlide.insertSheetsChartAsImage(chart2, leftPosition2, topPosition, desiredChartWidth, calculatedChartHeight);
                Logger.log(`      Embedded Chart 2 successfully.`);
             } catch(chartErr) {
                  Logger.log(`      ERROR embedding Chart 2: ${chartErr.message}`);
             }
        } else {
            Logger.log(`   WARNING: Chart 2 "${PIE_CHART_2_TITLE}" not found on sheet "${CHART_SHEET_NAME}".`);
        }
        // --- 7. Save and Close ---
        Logger.log("   Saving presentation...");
        presentation.saveAndClose();

        // --- 8. Optional: Add Link Back to Sheet ---
        try {
            const linkCell = 'A124'; // <<< Use your desired cell
            const execSummarySheet = ss.getSheetByName('Executive Summary'); // Get sheet object
            if (execSummarySheet) { // Check if sheet exists
                  execSummarySheet.getRange(linkCell).insertHyperlink(newPresentationUrl, newFileName);
                  Logger.log(`   Added link to presentation in Executive Summary!${linkCell}`);
            } else {
                  Logger.log(`   Warning: Could not find 'Executive Summary' sheet to insert link.`);
            }
        } catch (linkErr) {
            Logger.log(`   Warning: Could not insert presentation link into sheet: ${linkErr.message}`)
        }

        Logger.log(`[${functionName}] Presentation generation complete! URL: ${newPresentationUrl}`);
        ui.alert(`Presentation generated successfully!\n\nLink: ${newPresentationUrl}`);

    } catch (error) {
        Logger.log(`[${functionName}] ERROR: ${error.message} - Stack: ${error.stack || 'N/A'}`);
        ui.alert(`Error generating presentation: ${error.message}`);
    }
}
/**
 * Iteratively adjusts the Investor Equity Split (%) in a target cell
 * until the Investor Target IRR (%) in another cell reaches a desired value.
 *
 * @param {Sheet} sheet The 'Detailed Analysis' sheet object.
 * @param {number} targetIRR The desired target IRR (e.g., 0.25 for 25%).
 * @param {string} investorSplitCellA1 A1 notation of the cell containing the Investor Equity Split %.
 * @param {string} investorIRRCellA1 A1 notation of the cell containing the calculated Investor IRR %.
 * @param {string} projectNetProfitCellA1 A1 notation for the project net profit cell.
 * @returns {number | null} The final calculated investor split % or null if target not met.
 */
function calculateInvestorSplitForTargetIRR(sheet, targetIRR, investorSplitCellA1, investorIRRCellA1, projectNetProfitCellA1) {
    const functionName = "calculateInvestorSplitForTargetIRR";
    Logger.log(`[${functionName}] Attempting to find investor split for Target IRR: ${(targetIRR * 100).toFixed(2)}%`);

    const MAX_ITERATIONS = 100; // Safety limit to prevent infinite loops
    const TOLERANCE = 0.001; // Target IRR within +/- 0.1%
    let adjustmentStep = 0.01; // Start by adjusting 1% per step

    try {
        // Check if project is profitable first
        const projectNetProfit = sheet.getRange(projectNetProfitCellA1).getValue();
        if (typeof projectNetProfit !== 'number' || projectNetProfit <= 0) {
            Logger.log(`[${functionName}] Project Net Profit (${projectNetProfit}) is not positive. Cannot achieve target IRR. Skipping calculation.`);
            // Optional: Set a default split? Or just leave it as is?
            // sheet.getRange(investorSplitCellA1).setValue(0.5); // Example: Default to 50% if not profitable
            // SpreadsheetApp.flush();
            return null;
        }

        let currentInvestorSplit = sheet.getRange(investorSplitCellA1).getValue();
        // If cell is blank or invalid, start at 50%
        if (typeof currentInvestorSplit !== 'number' || isNaN(currentInvestorSplit) || currentInvestorSplit <=0 || currentInvestorSplit >=1) {
            currentInvestorSplit = 0.50; // Start guess at 50%
            Logger.log(`   Initial investor split invalid, starting guess at 50%`);
        } else {
             Logger.log(`   Starting with current investor split: ${(currentInvestorSplit * 100).toFixed(1)}%`);
        }


        for (let i = 0; i < MAX_ITERATIONS; i++) {
            // 1. Set the current guess
            sheet.getRange(investorSplitCellA1).setValue(currentInvestorSplit);
            SpreadsheetApp.flush(); // Force recalculation
            Utilities.sleep(750); // Pause longer for potentially complex IRR calcs

            // 2. Read the resulting IRR
            const currentIRR = sheet.getRange(investorIRRCellA1).getValue();
            Logger.log(`   Iteration ${i+1}: Split=${(currentInvestorSplit*100).toFixed(2)}% -> IRR=${(typeof currentIRR === 'number' ? (currentIRR*100).toFixed(2) : 'N/A')}%`);

            // Check if IRR calculation failed
            if (typeof currentIRR !== 'number' || isNaN(currentIRR)) {
                Logger.log(`   IRR calculation failed for current split. Trying slight adjustment.`);
                // Make a small adjustment and hope it recovers, or could stop here.
                 currentInvestorSplit += (adjustmentStep * 0.1); // Tiny nudge
                 currentInvestorSplit = Math.max(0.01, Math.min(0.99, currentInvestorSplit)); // Keep within bounds
                 continue; // Try next iteration
            }

            // 3. Check if target is met
            const difference = currentIRR - targetIRR;
            if (Math.abs(difference) <= TOLERANCE) {
                Logger.log(`[${functionName}] Target IRR reached! Final Investor Split: ${(currentInvestorSplit * 100).toFixed(2)}%`);
                return currentInvestorSplit; // Success!
            }

            // 4. Adjust the split for the next iteration
            if (difference > 0) {
                // Current IRR is TOO HIGH (investor getting too much) -> DECREASE investor split
                 Logger.log(`      IRR too high (${(difference*100).toFixed(2)}% diff). Decreasing investor split.`);
                 currentInvestorSplit -= adjustmentStep;
            } else {
                // Current IRR is TOO LOW -> INCREASE investor split
                 Logger.log(`      IRR too low (${(difference*100).toFixed(2)}% diff). Increasing investor split.`);
                 currentInvestorSplit += adjustmentStep;
            }

            // Ensure split stays within reasonable bounds (e.g., 1% to 99%)
             currentInvestorSplit = Math.max(0.01, Math.min(0.99, currentInvestorSplit));

             // Optional: Reduce step size as we get closer (finer tuning)
            if (Math.abs(difference) < 0.05) { // If within 5%
                 adjustmentStep = 0.005; // Use smaller steps
            } else {
                 adjustmentStep = 0.01; // Reset to larger steps if far away
            }

        } // End loop

        Logger.log(`[${functionName}] Maximum iterations (${MAX_ITERATIONS}) reached. Target IRR might not be achievable or calculation is unstable. Final Split: ${(currentInvestorSplit * 100).toFixed(2)}%`);
        return currentInvestorSplit; // Return the last attempted split

    } catch (error) {
        Logger.log(`[${functionName}] ERROR during calculation: ${error.message} - Stack: ${error.stack || 'N/A'}`);
        return null; // Return null on error
    }
}

// --- Debug Logging Helper (set Script Property DEBUG_LOG='true' to enable selective extra logs) ---
const DEBUG = getProp('DEBUG_LOG','false') === 'true';
function debugLog(msg){ if (DEBUG) Logger.log('[DEBUG] '+ msg); }