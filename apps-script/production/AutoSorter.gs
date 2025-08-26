// AutoSorter production copy
function autoSortSheet(){
	var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	var lastRow=sheet.getLastRow(); var lastColumn=sheet.getLastColumn(); var headerRows=3;
	if(lastRow>headerRows && lastColumn>0){ var dataRowCount=lastRow-headerRows; if(lastColumn<28){ Logger.log('Warning: helper col AB may be outside range.'); }
		var range=sheet.getRange(headerRows+1,1,dataRowCount,lastColumn);
		range.sort([{column:28,ascending:true},{column:24,ascending:false},{column:18,ascending:false}]);
		SpreadsheetApp.flush();
	} else { Logger.log('Not enough data rows to sort.'); }
}
