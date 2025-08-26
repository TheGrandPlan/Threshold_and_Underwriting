// =============================================================
// TACS (Threshold & Creation System)
// Production copy of threshold script.
// =============================================================
var THRESHOLD_SHEET_NAME = 'Deal Analysis Summary - Prospective Properties';
var HEADER_ROWS = 3;
var COL = { address:1, zip:2, askingPrice:3, lotSize:4, status:6, metric:17, processedCheck:20, link:21 };
function getProp_(key, fallback){ try{ var v=PropertiesService.getScriptProperties().getProperty(key); return v? v: fallback; }catch(e){ return fallback; }}
var PARENT_FOLDER_ID = getProp_('PARENT_FOLDER_ID','MISSING_PARENT_FOLDER_ID');
var TEMPLATE_FILE_ID = getProp_('TEMPLATE_FILE_ID','MISSING_TEMPLATE_FILE_ID');
var N8N_WEBHOOK_URL = getProp_('N8N_WEBHOOK_URL','');
var CENTRAL_URL = getProp_('CENTRAL_URL','');
function checkThresholdAndProcess(){
	var ss=SpreadsheetApp.getActiveSpreadsheet();
	var sheet=ss.getSheetByName(THRESHOLD_SHEET_NAME); if(!sheet){ Logger.log('Sheet missing'); return; }
	var data=sheet.getDataRange().getValues(); if(data.length<=HEADER_ROWS){ Logger.log('No rows'); return; }
	var processed=0; for(var i=HEADER_ROWS;i<data.length;i++){ try{ var row=data[i]; var metric=parseMetric_(row[COL.metric]); if(metric==null||!(metric>0.40)) continue; var status=String(row[COL.status]||'').trim().toUpperCase(); if(status!=='ACTIVE') continue; if(isChecked_(row[COL.processedCheck])) continue; var address=row[COL.address]; if(!address){ continue; } createFolderCopyAndCallCentralized(address,row[COL.zip],i,row[COL.lotSize],row[COL.askingPrice]); processed++; }catch(err){ Logger.log('Err row '+(i+1)+': '+err.message); }}
	Logger.log('Processed '+processed+' row(s).');
}
function createFolderCopyAndCallCentralized(propertyAddress,zipCode,rowIndex,lotSize,askingPrice){ var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(THRESHOLD_SHEET_NAME); var sheetRow=rowIndex+1; var runId=Utilities.getUuid().slice(0,8); try{ var parent=DriveApp.getFolderById(PARENT_FOLDER_ID); var safeName=sanitizeName_(propertyAddress); var folder=parent.createFolder(safeName); var tmpl=DriveApp.getFileById(TEMPLATE_FILE_ID); var newName=safeName+' - Analysis'; var file=tmpl.makeCopy(newName,folder); var url=file.getUrl(); sheet.getRange(sheetRow,COL.processedCheck+1).setValue(true); sheet.getRange(sheetRow,COL.link+1).setValue(url); populateNewSpreadsheet_(file.getId(),propertyAddress,zipCode,lotSize,askingPrice); if(N8N_WEBHOOK_URL){ postJsonSafe_(N8N_WEBHOOK_URL,{fileName:newName,fileUrl:url},'n8n',runId);} if(CENTRAL_URL){ var secret=Utilities.getUuid(); postJsonSafe_(CENTRAL_URL,{action:'initialize',spreadsheetId:file.getId(),callbackSecret:secret,propertyAddress:propertyAddress},'central',runId);} }catch(e){ Logger.log('[TACS '+runId+'] ERROR '+e.message); }}
function parseMetric_(raw){ if(raw==null||raw==='') return null; if(typeof raw==='number') return raw>1? raw/100: raw; var s=String(raw).trim(); if(!s) return null; var pct=s.indexOf('%')!==-1; var num=parseFloat(s.replace(/[^0-9.+-]/g,'')); if(isNaN(num)) return null; return (pct||num>1)? num/100: num; }
function isChecked_(v){ return v===true||v===1||String(v).toUpperCase()==='TRUE'; }
function sanitizeName_(n){ return String(n).replace(/[\\/:*?"<>|]+/g,' ').replace(/\s+/g,' ').trim().substring(0,150); }
function populateNewSpreadsheet_(fileId,address,zip,lotSize,askingPrice){ var ss=SpreadsheetApp.openById(fileId); var a=ss.getSheetByName('Detailed Analysis'); var area=ss.getSheetByName('Area Summary'); if(a){ a.getRange('B4').setValue(address+', Austin, TX '+(zip||'')); if(zip!=null&&String(zip).trim()!=='') a.getRange('B5').setValue(zip); a.getRange('B6').setValue(address); if(askingPrice!=null&&String(askingPrice).trim()!=='') a.getRange('B59').setValue(askingPrice);} if(area){ if(lotSize!=null&&String(lotSize).trim()!=='') area.getRange('B3').setValue(lotSize);} SpreadsheetApp.flush(); }
function postJsonSafe_(url,payload,label,runId){ try{ var resp=UrlFetchApp.fetch(url,{method:'post',contentType:'application/json',muteHttpExceptions:true,payload:JSON.stringify(payload)}); var code=resp.getResponseCode(); if(code>=400){ Logger.log('['+runId+'] '+label+' HTTP '+code+' body='+resp.getContentText().slice(0,200)); } }catch(e){ Logger.log('['+runId+'] '+label+' error '+e.message); }}

