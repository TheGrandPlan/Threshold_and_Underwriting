/**
 * Auto-sorts the active sheet based on specified column priorities.
 * Primary: Status Priority (Helper Column AB Ascending)
 * Secondary: Column X Descending
 * Tertiary: Column R Descending
 * Assumes data starts on row 4.
 */
function autoSortSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); // Get the active sheet
  var lastRow = sheet.getLastRow();  // Get the last row with data
  var lastColumn = sheet.getLastColumn();  // Get the last column with data
  var headerRows = 3; // Number of header rows to exclude

  // Ensure the sheet has data rows below the header
  if (lastRow > headerRows && lastColumn > 0) {
    // Calculate the range of data rows (start row 4)
    var dataRowCount = lastRow - headerRows;
    // Make sure the range includes up to at least column AB (or your helper column)
    if (lastColumn < 28) { // AB = 28
        Logger.log("Warning: Sorting range might not include the helper column AB. Adjust lastColumn if needed.");
        // Optionally force lastColumn to be at least 28 if you know it exists:
        // lastColumn = Math.max(lastColumn, 28);
    }
    var range = sheet.getRange(headerRows + 1, 1, dataRowCount, lastColumn);

    // Define sort order
    const HELPER_COLUMN_INDEX = 28; // <<< Column AB (Status Priority: 1, 2, 3, 4, 5, 99)
    const ROI_COLUMN_INDEX = 24;      // <<< Column X (ROI?)
    const METRIC_COLUMN_INDEX = 18;   // <<< Column R (Metric?)

    Logger.log(`Sorting range A4:${String.fromCharCode(64 + lastColumn)}${lastRow}...`);
    Logger.log(` Primary: Col ${HELPER_COLUMN_INDEX} (AB) Ascending (Status Priority)`);
    Logger.log(` Secondary: Col ${ROI_COLUMN_INDEX} (X) Descending`);
    Logger.log(` Tertiary: Col ${METRIC_COLUMN_INDEX} (R) Descending`);

    // Perform the sort based on the new priority:
    range.sort([
    {column: 28, ascending: true},    // PRIMARY: Status Priority (AB, Ascending)
    {column: 24, ascending: false},   // SECONDARY: ROI? (X, Descending)
    {column: 18, ascending: false}    // TERTIARY: Metric? (R, Descending)
  ]);

     Logger.log("Sort complete.");
     SpreadsheetApp.flush(); // Optional: attempt to apply sort visually immediately

  } else {
    Logger.log("Not enough data rows to sort (or no columns).");
  }
}