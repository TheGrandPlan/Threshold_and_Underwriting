function removeDuplicateAddresses() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange("B4:B").getValues();
  const seen = {};
  for (let i = data.length - 1; i >= 0; i--) {
    const address = data[i][0];
    if (seen[address]) {
      sheet.deleteRow(i + 3); // Adjust for header offset
    } else {
      seen[address] = true;
    }
  }
}
