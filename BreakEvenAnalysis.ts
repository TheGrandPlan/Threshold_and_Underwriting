
// --- UI CELL DEFINITIONS for "Scenario Distribution" Sheet ---
const SCENARIO_SHEET_NAME = "Scenario Distribution";
const UI_STATUS_LOG_CELL = "I8";

// Construction Cost
const COST_CHECKBOX_CELL = "I4";
const COST_DESIRED_INPUT_CELL = "J6";
const COST_CURRENT_VALUE_CELL = "I6"; // Reads from model
const COST_CALCULATED_BREAKEVEN_CELL = "K6"; // Script writes result here
const COST_PERCENT_CHANGE_CELL = "L6"; // Sheet formula: =(K6-I6)/I6 or =(J6-I6)/I6

// Sales Price
const PRICE_CHECKBOX_CELL = "M4";
const PRICE_DESIRED_INPUT_CELL = "N6";
const PRICE_CURRENT_VALUE_CELL = "M6"; // Reads from model
const PRICE_CALCULATED_BREAKEVEN_CELL = "O6"; // Script writes result here
const PRICE_PERCENT_CHANGE_CELL = "P6"; // Sheet formula

// Timeline (Months)
const TIMELINE_CHECKBOX_CELL = "Q4";
const TIMELINE_DESIRED_INPUT_CELL = "R6";
const TIMELINE_CURRENT_VALUE_CELL = "Q6"; // Reads from model
const TIMELINE_CALCULATED_BREAKEVEN_CELL = "S6"; // Script writes result here
const TIMELINE_PERCENT_CHANGE_CELL = "T6"; // Sheet formula

// Reset Button (concept, actual button will call a function)
const RESET_BUTTON_LOGIC_CELL = "N7"; // Cell where "Reset Data" text is

// --- CORE MODEL INPUT CELL DEFINITIONS (on 'Detailed Analysis' sheet) ---
const DA_SHEET_NAME = "Detailed Analysis";
const DA_CONSTRUCTION_COST_INPUT_CELL = "B82";
const DA_SALES_PRICE_INPUT_CELL = "B56";
const DA_TIMELINE_INPUT_CELL = "B32"; // Assuming this is in months

// --- TARGET OUTPUT CELL for Break-even ---
const DA_PROJECT_NET_PROFIT_OUTPUT_CELL = "B137"; // We aim to make this $0

// --- Constants for Iterative Calculation ---
const BREAKEVEN_MAX_ITERATIONS = 200;
const BREAKEVEN_PROFIT_TOLERANCE = 100; // Aim for profit within +/- $100 of zero

/**
 * Wrapper function for Break-Even Analysis.
 * Reads UI settings, updates model with fixed variables, runs iterative calculation.
 * Leaves model in the state that achieved break-even.
 * Calculates and writes the "% of change" for ALL variables based on their final state vs baseline.
 * Original values are NOT automatically restored here; use Reset button for that.
 */
function runBreakevenAnalysis() {
    const functionName = "runBreakevenAnalysis";
    updateStatusCell("Starting Break-Even Analysis...");
    Logger.log(`[${functionName}] Starting...`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scenarioSheet = ss.getSheetByName(SCENARIO_SHEET_NAME);
    const detailedAnalysisSheet = ss.getSheetByName(DA_SHEET_NAME);
    let ui;

    if (!scenarioSheet || !detailedAnalysisSheet) {
        // ... (error handling as before) ...
        return;
    }

    updateStatusCell("Identifying variable to solve for...");
    let solveForVariable = null;
    let variableToSolveCell_DA = null;
    let resultCell_Scenario = null;
    let variableFriendlyName = "";

    // --- Read initial baseline values from Detailed Analysis for % change calculation later ---
    const baselineValues = {
        cost: detailedAnalysisSheet.getRange(DA_CONSTRUCTION_COST_INPUT_CELL).getValue(),
        price: detailedAnalysisSheet.getRange(DA_SALES_PRICE_INPUT_CELL).getValue(),
        timeline: detailedAnalysisSheet.getRange(DA_TIMELINE_INPUT_CELL).getValue()
    };
    Logger.log(`   Baseline Model Values from DA: Cost=${baselineValues.cost}, Price=${baselineValues.price}, Timeline=${baselineValues.timeline}`);

    if (scenarioSheet.getRange(COST_CHECKBOX_CELL).isChecked()) {
        solveForVariable = "ConstructionCost"; variableFriendlyName = "Max Construction Cost";
        variableToSolveCell_DA = DA_CONSTRUCTION_COST_INPUT_CELL; resultCell_Scenario = COST_CALCULATED_BREAKEVEN_CELL;
    } else if (scenarioSheet.getRange(PRICE_CHECKBOX_CELL).isChecked()) {
        solveForVariable = "SalesPrice"; variableFriendlyName = "Min Sales Price";
        variableToSolveCell_DA = DA_SALES_PRICE_INPUT_CELL; resultCell_Scenario = PRICE_CALCULATED_BREAKEVEN_CELL;
    } else if (scenarioSheet.getRange(TIMELINE_CHECKBOX_CELL).isChecked()) {
        solveForVariable = "Timeline"; variableFriendlyName = "Max Timeline";
        variableToSolveCell_DA = DA_TIMELINE_INPUT_CELL; resultCell_Scenario = TIMELINE_CALCULATED_BREAKEVEN_CELL;
    }

    if (!solveForVariable) {
        // ... (error handling as before) ...
        return;
    }
    updateStatusCell(`Solving for: ${variableFriendlyName}...`);
    Logger.log(`   Solving for: ${solveForVariable}`);
    SpreadsheetApp.flush();

    try {
        updateStatusCell(`Applying fixed desired inputs...`);
        Logger.log("   Applying desired inputs for fixed variables (as values)...");

        // Apply desired inputs and store what value was actually set in DA for fixed variables
        const finalModelInputs = { cost: baselineValues.cost, price: baselineValues.price, timeline: baselineValues.timeline };

        if (solveForVariable !== "ConstructionCost") {
            const desiredCost = scenarioSheet.getRange(COST_DESIRED_INPUT_CELL).getValue();
            if (desiredCost !== "" && typeof desiredCost === 'number') {
                detailedAnalysisSheet.getRange(DA_CONSTRUCTION_COST_INPUT_CELL).setValue(desiredCost);
                finalModelInputs.cost = desiredCost; // Store the value that was set
                Logger.log(`      Set Construction Cost to (temp value): ${desiredCost}`);
            } else { // If desired is blank, model uses its baseline (which we already set DA to)
                detailedAnalysisSheet.getRange(DA_CONSTRUCTION_COST_INPUT_CELL).setValue(baselineValues.cost);
                finalModelInputs.cost = baselineValues.cost;
            }
        }
        if (solveForVariable !== "SalesPrice") {
            const desiredPrice = scenarioSheet.getRange(PRICE_DESIRED_INPUT_CELL).getValue();
            if (desiredPrice !== "" && typeof desiredPrice === 'number') {
                detailedAnalysisSheet.getRange(DA_SALES_PRICE_INPUT_CELL).setValue(desiredPrice);
                finalModelInputs.price = desiredPrice;
                Logger.log(`      Set Sales Price to (temp value): ${desiredPrice}`);
            } else {
                detailedAnalysisSheet.getRange(DA_SALES_PRICE_INPUT_CELL).setValue(baselineValues.price);
                finalModelInputs.price = baselineValues.price;
            }
        }
        if (solveForVariable !== "Timeline") {
            const desiredTimeline = scenarioSheet.getRange(TIMELINE_DESIRED_INPUT_CELL).getValue();
            if (desiredTimeline !== "" && typeof desiredTimeline === 'number') {
                detailedAnalysisSheet.getRange(DA_TIMELINE_INPUT_CELL).setValue(desiredTimeline);
                finalModelInputs.timeline = desiredTimeline;
                Logger.log(`      Set Timeline to (temp value): ${desiredTimeline}`);
            } else {
                detailedAnalysisSheet.getRange(DA_TIMELINE_INPUT_CELL).setValue(baselineValues.timeline);
                finalModelInputs.timeline = baselineValues.timeline;
            }
        }
        SpreadsheetApp.flush();
        Utilities.sleep(500);

        updateStatusCell(`Calculating break-even for ${variableFriendlyName}... (Iter 0)`);
        SpreadsheetApp.getUi().alert("Calculating break-even... please wait.\nClick OK to proceed.\n(Status will update in cell " + UI_STATUS_LOG_CELL + ")");

        let initialDirection, initialStep;
        const currentValueToSolve = detailedAnalysisSheet.getRange(variableToSolveCell_DA).getValue();

        if (solveForVariable === "ConstructionCost") { initialDirection = 1; initialStep = currentValueToSolve * 0.01 || 10000; }
        else if (solveForVariable === "SalesPrice") { initialDirection = -1; initialStep = currentValueToSolve * 0.01 || 10000; }
        else { initialDirection = 1; initialStep = 1; }

        const breakevenValue = calculateBreakevenForSingleVariable(
            detailedAnalysisSheet, variableToSolveCell_DA, DA_PROJECT_NET_PROFIT_OUTPUT_CELL,
            0, initialDirection, initialStep, BREAKEVEN_PROFIT_TOLERANCE, BREAKEVEN_MAX_ITERATIONS
        );

        // Update finalModelInputs with the solved-for variable's breakeven value
        if (breakevenValue !== null) {
            if (solveForVariable === "ConstructionCost") finalModelInputs.cost = breakevenValue;
            else if (solveForVariable === "SalesPrice") finalModelInputs.price = breakevenValue;
            else if (solveForVariable === "Timeline") finalModelInputs.timeline = breakevenValue;
        } else { // If breakeven failed, use the last attempted value for the solved variable
            if (solveForVariable === "ConstructionCost") finalModelInputs.cost = detailedAnalysisSheet.getRange(DA_CONSTRUCTION_COST_INPUT_CELL).getValue();
            else if (solveForVariable === "SalesPrice") finalModelInputs.price = detailedAnalysisSheet.getRange(DA_SALES_PRICE_INPUT_CELL).getValue();
            else if (solveForVariable === "Timeline") finalModelInputs.timeline = detailedAnalysisSheet.getRange(DA_TIMELINE_INPUT_CELL).getValue();
        }


        // --- Calculate and Set ALL % of Change values ---
        Logger.log("   Calculating and setting all % of change values...");
        // Cost % Change
        if (typeof baselineValues.cost === 'number' && baselineValues.cost !== 0) {
            const costChange = (finalModelInputs.cost - baselineValues.cost) / baselineValues.cost;
            scenarioSheet.getRange(COST_PERCENT_CHANGE_CELL).setValue(costChange).setNumberFormat("0.00%");
        } else if (finalModelInputs.cost !== baselineValues.cost) { // If baseline was 0 but value changed
            scenarioSheet.getRange(COST_PERCENT_CHANGE_CELL).setValue("N/A (from 0)");
        } else { scenarioSheet.getRange(COST_PERCENT_CHANGE_CELL).clearContent(); }

        // Price % Change
        if (typeof baselineValues.price === 'number' && baselineValues.price !== 0) {
            const priceChange = (finalModelInputs.price - baselineValues.price) / baselineValues.price;
            scenarioSheet.getRange(PRICE_PERCENT_CHANGE_CELL).setValue(priceChange).setNumberFormat("0.00%");
        } else if (finalModelInputs.price !== baselineValues.price) {
            scenarioSheet.getRange(PRICE_PERCENT_CHANGE_CELL).setValue("N/A (from 0)");
        } else { scenarioSheet.getRange(PRICE_PERCENT_CHANGE_CELL).clearContent(); }

        // Timeline % Change
        if (typeof baselineValues.timeline === 'number' && baselineValues.timeline !== 0) {
            const timelineChange = (finalModelInputs.timeline - baselineValues.timeline) / baselineValues.timeline;
            scenarioSheet.getRange(TIMELINE_PERCENT_CHANGE_CELL).setValue(timelineChange).setNumberFormat("0.00%");
        } else if (finalModelInputs.timeline !== baselineValues.timeline) {
            scenarioSheet.getRange(TIMELINE_PERCENT_CHANGE_CELL).setValue("N/A (from 0)");
        } else { scenarioSheet.getRange(TIMELINE_PERCENT_CHANGE_CELL).clearContent(); }
        SpreadsheetApp.flush();


        ui = SpreadsheetApp.getUi();
        if (breakevenValue !== null) {
            scenarioSheet.getRange(resultCell_Scenario).setValue(breakevenValue);
            const currentNetProfit = detailedAnalysisSheet.getRange(DA_PROJECT_NET_PROFIT_OUTPUT_CELL).getDisplayValue();
            const resultMsg = `Break-Even for ${variableFriendlyName}: ${breakevenValue.toFixed(solveForVariable === "Timeline" ? 1 : 0)}. Final Net Profit: ${currentNetProfit}`;
            ui.alert(`Break-Even Analysis Complete!\n\n${resultMsg}\n\nModel inputs in 'Detailed Analysis' are now set to this break-even scenario.\nUse 'Reset Data' to restore original formulas.`);
            updateStatusCell(resultMsg);
            Logger.log(`   ${resultMsg}`);
        } else {
            const failMsg = `Could not find break-even for ${variableFriendlyName}. Check logs.`;
            ui.alert(failMsg + " Model inputs in 'Detailed Analysis' may have been altered.");
            updateStatusCell(failMsg);
            Logger.log(`   ${failMsg}`);
        }
        // Removed final toast

    } catch (error) {
        const errorMsg = `Error: ${error.message}`;
        Logger.log(`[${functionName}] ERROR: ${errorMsg} - Stack: ${error.stack || 'N/A'}`);
        ui = SpreadsheetApp.getUi();
        ui.alert(`An error occurred: ${errorMsg}`);
        updateStatusCell(`Error during analysis. Check logs.`);
    }

    Logger.log(`[${functionName}] Break-Even Analysis Attempt Finished. Model reflects last calculated state.`);
    updateStatusCell("Break-Even analysis finished. Press 'Reset Data' to restore default model formulas.");
}
/**
 * Iteratively changes a single input variable to find the point where a target output cell reaches a target value.
 *
 * @param {Sheet} sheet The sheet object where inputs and outputs reside (e.g., 'Detailed Analysis').
 * @param {string} variableCellA1 A1 notation of the input cell to change.
 * @param {string} targetOutputCellA1 A1 notation of the output cell to monitor.
 * @param {number} targetOutputValue The desired value for the output cell (e.g., 0 for breakeven profit).
 * @param {number} initialDirection 1 to increase the variable, -1 to decrease it.
 * @param {number} initialStep The initial amount to change the variable by.
 * @param {number} tolerance How close the output needs to be to the target value.
 * @param {number} maxIterations Safety limit for iterations.
 * @returns {number | null} The value of the variableCellA1 at breakeven, or null if not found.
 * Iteratively changes a single input variable to find the point where a target output cell reaches a target value.
 * Updates a UI status cell during iterations.
 */
function calculateBreakevenForSingleVariable(sheet, variableCellA1, targetOutputCellA1, targetOutputValue, initialDirection, initialStep, tolerance, maxIterations) {
    const functionName = "calculateBreakevenForSingleVariable";
    updateStatusCell(`Starting iterations for ${variableCellA1}...`); // Initial status for this function
    Logger.log(`[${functionName}] Solving for ${variableCellA1} to make ${targetOutputCellA1} = ${targetOutputValue} (Step: ${initialStep}, Dir: ${initialDirection})`);

    let currentVariableValue = sheet.getRange(variableCellA1).getValue();
    if (typeof currentVariableValue !== 'number') {
        const errMsg = `Initial value in ${variableCellA1} is not a number. Aborting.`;
        Logger.log(`   ${errMsg}`);
        updateStatusCell(errMsg);
        return null;
    }

    let step = initialStep;
    let direction = initialDirection;
    let lastDifference = null;

    for (let i = 0; i < maxIterations; i++) {
        sheet.getRange(variableCellA1).setValue(currentVariableValue);
        SpreadsheetApp.flush();
        Utilities.sleep(750); // Allow time for sheet recalculation

        const currentOutputValue = sheet.getRange(targetOutputCellA1).getValue();
        if (typeof currentOutputValue !== 'number' || isNaN(currentOutputValue)) {
            const errMsg = `Iter ${i+1}: Output in ${targetOutputCellA1} is NAN ('${currentOutputValue}'). Var: ${currentVariableValue}. Stopping.`;
            Logger.log(`   ${errMsg}`);
            updateStatusCell(errMsg);
            return null;
        }

        const difference = currentOutputValue - targetOutputValue;
        const statusMsg = `Iter ${i+1}: ${variableCellA1}=${currentVariableValue.toFixed(2)}, Profit=${currentOutputValue.toLocaleString('en-US', { style: 'currency', currency: 'USD' })}`;
        Logger.log(`   ${statusMsg}, Diff=${difference.toFixed(2)}`);
        updateStatusCell(statusMsg); // Update UI cell with iteration status

        if (Math.abs(difference) <= tolerance) {
            Logger.log(`[${functionName}] Target reached! Breakeven for ${variableCellA1} is ${currentVariableValue.toFixed(4)}`);
            // updateStatusCell already called with final values before this return
            return currentVariableValue;
        }

        const previousVariableValue = currentVariableValue; // Store before adjusting

        if (lastDifference !== null && Math.sign(difference) !== Math.sign(lastDifference) && Math.abs(difference) > tolerance) {
            direction *= -1;
            step /= 2;
            Logger.log(`      Overshot. Reversed direction. New step: ${step.toFixed(6)}`);
            updateStatusCell(`Overshot. Reducing step for ${variableCellA1}.`);
        }

        currentVariableValue += (step * direction);
        lastDifference = difference;

        // Sanity checks
        if (variableCellA1 === DA_TIMELINE_INPUT_CELL && currentVariableValue < 1) currentVariableValue = 1;
        if ((variableCellA1 === DA_CONSTRUCTION_COST_INPUT_CELL || variableCellA1 === DA_SALES_PRICE_INPUT_CELL) && currentVariableValue < 0) {
            const errMsg = `Variable ${variableCellA1} went negative. Stopping.`;
            Logger.log(`   ${errMsg}`);
            updateStatusCell(errMsg);
            return previousVariableValue; // Return last valid value
        }
         if (step < (variableCellA1 === DA_TIMELINE_INPUT_CELL ? 0.5 : 0.01) ) { // Adjusted minimum step
             Logger.log(`   Step size too small for ${variableCellA1}. Converged or stuck at ${currentVariableValue.toFixed(4)}.`);
             updateStatusCell(`Converged/stuck for ${variableCellA1} at ${currentVariableValue.toFixed(2)}`);
             return currentVariableValue;
         }
         if (currentVariableValue === previousVariableValue && Math.abs(difference) > tolerance){
             Logger.log(`   Variable ${variableCellA1} value stuck at ${currentVariableValue.toFixed(4)} but tolerance not met. Exiting.`);
             updateStatusCell(`Value stuck for ${variableCellA1}. Tolerance not met.`);
             return currentVariableValue;
         }
    }

    const finalMsg = `Max iterations for ${variableCellA1}. Target not met. Last value: ${currentVariableValue.toFixed(4)}`;
    Logger.log(`[${functionName}] ${finalMsg}`);
    updateStatusCell(finalMsg);
    return currentVariableValue;
}

/**
 * Resets the UI on "Scenario Distribution" sheet and restores master formulas
 * to 'Detailed Analysis' core input cells.
 */
function resetBreakevenInputs() {
    const functionName = "resetBreakevenInputs";
    updateStatusCell("Resetting inputs and model formulas...");
    Logger.log(`[${functionName}] Resetting UI data and core model formulas...`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const scenarioSheet = ss.getSheetByName(SCENARIO_SHEET_NAME);
    const detailedAnalysisSheet = ss.getSheetByName(DA_SHEET_NAME);
    let ui;

    if (!scenarioSheet || !detailedAnalysisSheet) {
        const errMsg = "Error: Required sheets ('Scenario Distribution' or 'Detailed Analysis') not found for reset.";
        ui = SpreadsheetApp.getUi(); ui.alert(errMsg);
        Logger.log(`[${functionName}] ${errMsg}`); updateStatusCell(errMsg); return;
    }

    try {
        // Clear "Desired Value", "Calculated Breakeven", and "% of change" on Scenario Sheet
        scenarioSheet.getRange(COST_DESIRED_INPUT_CELL).clearContent();
        scenarioSheet.getRange(COST_CALCULATED_BREAKEVEN_CELL).clearContent();
        scenarioSheet.getRange(COST_PERCENT_CHANGE_CELL).clearContent(); // <<< ADDED

        scenarioSheet.getRange(PRICE_DESIRED_INPUT_CELL).clearContent();
        scenarioSheet.getRange(PRICE_CALCULATED_BREAKEVEN_CELL).clearContent();
        scenarioSheet.getRange(PRICE_PERCENT_CHANGE_CELL).clearContent(); // <<< ADDED

        scenarioSheet.getRange(TIMELINE_DESIRED_INPUT_CELL).clearContent();
        scenarioSheet.getRange(TIMELINE_CALCULATED_BREAKEVEN_CELL).clearContent();
        scenarioSheet.getRange(TIMELINE_PERCENT_CHANGE_CELL).clearContent(); // <<< ADDED

        scenarioSheet.getRange(UI_STATUS_LOG_CELL).clearContent();
        Logger.log("   Cleared Desired/Calculated/PercentChange/Status values on Scenario sheet.");

        // --- Re-apply MASTER FORMULAS (in A1 notation) to Detailed Analysis input cells ---
        Logger.log("   Resetting core model formulas on Detailed Analysis sheet...");
        const costFormulaA1 = "=$B$21*B80";
        const priceFormulaA1 = "=SUM(B53:B55)";
        const timelineFormulaA1 = "=SUM(B29:B31)";

        detailedAnalysisSheet.getRange(DA_CONSTRUCTION_COST_INPUT_CELL).setFormula(costFormulaA1);
        detailedAnalysisSheet.getRange(DA_SALES_PRICE_INPUT_CELL).setFormula(priceFormulaA1);
        detailedAnalysisSheet.getRange(DA_TIMELINE_INPUT_CELL).setFormula(timelineFormulaA1);
        Logger.log("      Restored master formulas to Detailed Analysis.");

        SpreadsheetApp.flush();
        updateStatusCell("UI inputs cleared. Model formulas reset to defaults.");
        ui = SpreadsheetApp.getUi(); ui.alert("Break-Even UI inputs cleared and core model formulas have been reset to their defaults.");
        Logger.log(`[${functionName}] Reset complete.`);

    } catch (error) {
        const errorMsg = `Error during reset: ${error.message}`;
        Logger.log(`[${functionName}] ERROR: ${errorMsg}`);
        updateStatusCell(errorMsg);
        ui = SpreadsheetApp.getUi(); ui.alert(errorMsg);
    }
}
/**
 * Updates a dedicated status cell on the UI sheet with a message.
 * @param {string} message The message to display.
 */
function updateStatusCell(message) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet(); // Assuming this script is bound to the sheet
        const scenarioSheet = ss.getSheetByName(SCENARIO_SHEET_NAME);
        if (scenarioSheet) {
            scenarioSheet.getRange(UI_STATUS_LOG_CELL).setValue(message);
            SpreadsheetApp.flush(); // Attempt to update the UI immediately
        }
    } catch (e) {
        Logger.log(`Error updating status cell (${UI_STATUS_LOG_CELL}): ${e.message}`);
    }
}
/**
 * Helper to attempt to translate an A1 formula to R1C1 for a specific target cell.
 * This is a workaround as Apps Script doesn't have a direct string-to-string A1->R1C1.
 * @param {string} a1Formula The formula in A1 notation.
 * @param {string} targetCellA1 The cell where this formula will be placed (e.g., "B82").
 * @return {string} The formula in R1C1 notation, or the original A1 if conversion fails.
 */
function translateToR1C1(a1Formula, targetCellA1) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        // Use a temporary sheet or a hidden helper cell if paranoid about active sheet
        const tempSheet = ss.insertSheet('TempFormulaConverterSheet');
        const tempCell = tempSheet.getRange("A1");
        tempCell.setFormula(a1Formula); // Set the A1 formula
        const r1c1Formula = tempCell.getFormulaR1C1(); // Get its R1C1 equivalent
        ss.deleteSheet(tempSheet); // Clean up
        return r1c1Formula || a1Formula; // Return R1C1 or original if R1C1 is blank
    } catch (e) {
        Logger.log(`Error translating formula "${a1Formula}" to R1C1 for target ${targetCellA1}: ${e.message}. Returning A1.`);
        return a1Formula; // Fallback to A1 if error
    }
}