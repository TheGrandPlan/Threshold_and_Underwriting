function geocodeAddresses() {
  Logger.log("Starting `geocodeAddresses` function...");

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    Logger.log(`Sheet "${SHEET_NAME}" not found.`);
    return;
  }
  Logger.log(`Sheet "${SHEET_NAME}" found successfully.`);

  // Retrieve last processed row from script properties or start from row 2
  let lastRowProcessed = parseInt(PropertiesService.getScriptProperties().getProperty('LAST_ROW')) || 2;
  const lastRow = sheet.getLastRow();
  const endRow = Math.min(lastRowProcessed + ROWS_PER_BATCH, lastRow);

  Logger.log(`Processing rows ${lastRowProcessed} to ${endRow}...`);

  // Retrieve address components
  const streetAddresses = sheet.getRange(`A${lastRowProcessed}:A${endRow}`).getValues();
  const cities = sheet.getRange(`C${lastRowProcessed}:C${endRow}`).getValues();
  const states = sheet.getRange(`D${lastRowProcessed}:D${endRow}`).getValues();
  const zips = sheet.getRange(`E${lastRowProcessed}:E${endRow}`).getValues();

  for (let i = 0; i < streetAddresses.length; i++) {
    const street = streetAddresses[i][0];
    const city = cities[i][0];
    const state = states[i][0];
    const zip = zips[i][0];

    // Combine the components into a full address
    const fullAddress = `${street}, ${city}, ${state} ${zip}`;
    Logger.log(`Full address for row ${lastRowProcessed + i}: ${fullAddress}`);
    
    if (!fullAddress || fullAddress.includes("undefined")) {
      Logger.log(`Skipping row ${lastRowProcessed + i} due to incomplete address.`);
      sheet.getRange(lastRowProcessed + i, 42).setValue("Incomplete Address");
      continue;
    }

    // Construct and send request for each address
    const url = `https://maps.googleapis.com/maps/api/geocode/json?address=${encodeURIComponent(fullAddress)}&key=${API_KEY}`;
    Logger.log(`Requesting URL: ${url}`);

    try {
      const response = UrlFetchApp.fetch(url);
      const data = JSON.parse(response.getContentText());
      
      if (data.status === "OK") {
        const location = data.results[0].geometry.location;
        Logger.log(`Coordinates for ${fullAddress}: ${location.lat}, ${location.lng}`);
        sheet.getRange(lastRowProcessed + i, 42).setValue(`${location.lat}, ${location.lng}`);
      } else {
        Logger.log(`No coordinates found for ${fullAddress}. Status: ${data.status}`);
        sheet.getRange(lastRowProcessed + i, 42).setValue("Not Found");
      }
    } catch (error) {
      Logger.log(`Error fetching coordinates for ${fullAddress}: ${error}`);
      sheet.getRange(lastRowProcessed + i, 42).setValue("Error");
    }
    Utilities.sleep(100);  // Delay to avoid rate limits
  }

  // Save the last processed row for the next execution
  PropertiesService.getScriptProperties().setProperty('LAST_ROW', endRow + 1);
  Logger.log(`Processing complete for rows ${lastRowProcessed} to ${endRow}. Next start row: ${endRow + 1}`);
}
