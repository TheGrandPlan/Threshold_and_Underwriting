// DeleteDuplicates production copy
function removeDuplicateAddresses(){
	var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	var data=sheet.getRange('B4:B').getValues();
	var seen={};
	for(var i=data.length-1;i>=0;i--){ var addr=data[i][0]; if(seen[addr]){ sheet.deleteRow(i+3); } else { seen[addr]=true; } }
}
