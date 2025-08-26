// =============================================================
// GeoCompEngine.gs (Production Copy)
//  - Compressed deployment version of local per-deal analysis engine
//  - Includes: setup/menu, locking, comps import/filtering, analysis outputs,
//    map + chart adjustments, preliminary sheet update, investor split optimizer.
//  - Break-even analysis lives in separate BreakEvenAnalysis.gs (already in production)
//  - Requires Script Properties (see README): PRELIMINARY_SHEET_ID, DATA_SPREADSHEET_ID,
//    SLIDES_TEMPLATE_ID, apiKey, staticMapsApiKey (+ credentials after setup)
//  - DEBUG_LOG='true' enables extra debug output
// =============================================================

function getProp(k,f){try{var v=PropertiesService.getScriptProperties().getProperty(k);return(v!==null&&v!==''?v:f);}catch(e){return f}}
// Helper to check if a string looks like a valid Drive/Spreadsheet ID (very loose heuristic)
function _looksLikeId(id){return id && id.length>20 && id.indexOf(' ')===-1 && id.indexOf('/')===-1 && !/^MISSING_/.test(id);} // heuristic only
// Safe spreadsheet opener with diagnostics; trims ID and logs length/first+last chars
function _openSpreadsheetByIdSafe(rawId,label){
	try{
		if(!rawId){Logger.log('[INIT] '+label+' ID missing');return null;}
		var id=String(rawId).trim();
		if(id!==rawId) Logger.log('[INIT] '+label+' ID had surrounding whitespace (trimmed).');
		Logger.log('[INIT] Opening '+label+' len='+id.length+' preview='+(id.substring(0,4)+'...'+id.substring(id.length-4)));
		var ss=SpreadsheetApp.openById(id); // will throw if invalid
		return ss;
	}catch(e){
		Logger.log('[INIT] ERROR opening '+label+' id="'+rawId+'" -> '+e);
		return null;
	}
}
function debugConfigStatus(){
	var sp=PropertiesService.getScriptProperties();
	var keys=['PRELIMINARY_SHEET_ID','DATA_SPREADSHEET_ID','SLIDES_TEMPLATE_ID','apiKey','staticMapsApiKey','SETUP_COMPLETE'];
	keys.forEach(function(k){ Logger.log('[CONFIG] '+k+'='+(sp.getProperty(k)||'(unset)')); });
}
const PRELIMINARY_SHEET_ID=getProp('PRELIMINARY_SHEET_ID','');
const PRELIMINARY_SHEET_NAME='Deal Analysis Summary - Prospective Properties';
const SLIDES_TEMPLATE_ID=getProp('SLIDES_TEMPLATE_ID','');
const CHART_SHEET_NAME='Executive Summary';
const PIE_CHART_1_TITLE='Project Timeline';
const PIE_CHART_2_TITLE='Investment vs. Profit';
const TARGET_IRR_INPUT_CELL='D153';
const INVESTOR_SPLIT_CELL='B145';
const INVESTOR_IRR_CELL='B153';
const PROJECT_NET_PROFIT_CELL='B137';
const METRIC_CELLS=Object.freeze({SIMPLE_ADDRESS:'B6',NET_PROFIT:'K109',ROI:'K111',MARGIN:'K110',ERROR_FEEDBACK:'A150'});
const DEBUG=getProp('DEBUG_LOG','true')==='true';function debugLog(m){if(DEBUG)Logger.log('[DEBUG] '+m)}
// Centralized color/style mapping for enforced scatter chart styling (Option C)
// Series index assumptions (both scatter charts):
//  0 = Subject (Project Sale Price)  -> Red
//  1 = Comparable Sales              -> Green
//  2 = Average/Median (summary cols) -> Purple
// If future chart definition changes (series re-ordered), update mapping or add detection.
const CHART_SERIES_STYLE={0:{color:'#d32f2f',pointSize:7},1:{color:'#2e7d32',pointSize:6},2:{color:'#6a1b9a',pointSize:6}}; // material-ish palette
// Optional: adjust legend behavior; set to 'none' to reduce clutter since subject is label-annotated already.
const CHART_LEGEND_POSITION='none';
function withGlobalLock(key,fn){var lock=LockService.getScriptLock();if(!lock.tryLock(5000)){Logger.log('[LOCK] Busy '+key);return;}try{return fn()}finally{lock.releaseLock()}}

const context={spreadsheet:null,sheet:null,dataSpreadsheet:null,dataSheet:null,compProperties:[],subjectCoords:null,config:{MASTER_SPREADSHEET_ID:SpreadsheetApp.getActiveSpreadsheet().getId(),MASTER_SHEET_NAME:'Sales Comps',ADDRESS_CELL:'A3',COMP_RADIUS_CELL:'P4',DATE_FILTER_CELL:'P5',AGE_FILTER_CELL:'P7',SIZE_FILTER_CELL:'P9',SUBJECT_SIZE_CELL:'G3',ANNUNCIATOR_CELL:'P10',SD_MULTIPLIER_CELL:'P11',COMP_RESULTS_START_ROW:33,COMP_RESULTS_START_COLUMN:'A',DATA_SPREADSHEET_ID:getProp('DATA_SPREADSHEET_ID',''),DATA_SHEET_NAME:'Current Comps'},filters:{radius:0,date:null,yearBuilt:null,sizePercentage:0},visibleRows:[],_visibleCompCache:null};
function initializeContext(){
	try{
		// Refresh data spreadsheet ID from properties if it was populated after script load
		try{var latestDataId=getProp('DATA_SPREADSHEET_ID','');if(_looksLikeId(latestDataId)&&latestDataId!==context.config.DATA_SPREADSHEET_ID){context.config.DATA_SPREADSHEET_ID=latestDataId;debugLog('initializeContext: refreshed DATA_SPREADSHEET_ID from properties.');}}catch(_r){}
		// Master spreadsheet: prefer active (allowed in simple triggers) then fallback to openById
		if(!context.spreadsheet){
			try{context.spreadsheet=SpreadsheetApp.getActiveSpreadsheet();Logger.log('[INIT] Obtained active spreadsheet directly.');}catch(_ae){}
			if(!context.spreadsheet){
				context.spreadsheet=_openSpreadsheetByIdSafe(context.config.MASTER_SPREADSHEET_ID,'MASTER_SPREADSHEET');
				if(!context.spreadsheet){Logger.log('[INIT] Aborting: master spreadsheet unavailable.');return;}
			}
			debugLog('initializeContext: master spreadsheet ready id='+context.config.MASTER_SPREADSHEET_ID);
		}
		if(!context.sheet){
			context.sheet=context.spreadsheet.getSheetByName(context.config.MASTER_SHEET_NAME);
			if(!context.sheet){Logger.log('[INIT] Master sheet '+context.config.MASTER_SHEET_NAME+' not found.');return;} else debugLog('initializeContext: master sheet located '+context.config.MASTER_SHEET_NAME);
		}
		// Data spreadsheet conditional
		if(!_looksLikeId(context.config.DATA_SPREADSHEET_ID)){
			Logger.log('[INIT] DATA_SPREADSHEET_ID missing/invalid heuristic; skip comps for now. Value='+(context.config.DATA_SPREADSHEET_ID||'(unset)'));
			return; // partial context allowed
		}
		if(!context.dataSpreadsheet){
			context.dataSpreadsheet=_openSpreadsheetByIdSafe(context.config.DATA_SPREADSHEET_ID,'DATA_SPREADSHEET');
			if(!context.dataSpreadsheet){Logger.log('[INIT] Could not open data spreadsheet; comps features disabled this run.');return;} else debugLog('initializeContext: data spreadsheet opened');
		}
		if(!context.dataSheet){
			context.dataSheet=context.dataSpreadsheet.getSheetByName(context.config.DATA_SHEET_NAME);
			if(!context.dataSheet){Logger.log('[INIT] Data sheet '+context.config.DATA_SHEET_NAME+' not found inside data spreadsheet.');} else debugLog('initializeContext: data sheet located '+context.config.DATA_SHEET_NAME+' rows='+context.dataSheet.getLastRow());
		}
	}catch(e){
		Logger.log('Context init error '+e);
	}
}

// One-off diagnostic you can run manually from the editor to inspect IDs explicitly
function debugTestIds(){
	var sp=PropertiesService.getScriptProperties();
	Logger.log('[TEST IDS] MASTER_SPREADSHEET_ID='+context.config.MASTER_SPREADSHEET_ID);
	Logger.log('[TEST IDS] DATA_SPREADSHEET_ID='+sp.getProperty('DATA_SPREADSHEET_ID'));
	Logger.log('[TEST IDS] PRELIMINARY_SHEET_ID='+sp.getProperty('PRELIMINARY_SHEET_ID'));
	_openSpreadsheetByIdSafe(context.config.MASTER_SPREADSHEET_ID,'MASTER_SPREADSHEET');
	_openSpreadsheetByIdSafe(sp.getProperty('DATA_SPREADSHEET_ID'),'DATA_SPREADSHEET');
	_openSpreadsheetByIdSafe(sp.getProperty('PRELIMINARY_SHEET_ID'),'PRELIMINARY_SPREADSHEET');
}

function onOpen(){
	var ui=SpreadsheetApp.getUi(),menu=ui.createMenu('⚙️ Setup');
	try{
		var sp=PropertiesService.getScriptProperties(),done=sp.getProperty('SETUP_COMPLETE');
		if(done!=='true'){
			menu.addItem('▶️ Run Initial Setup (Required Once)','runInitialSetup');
			SpreadsheetApp.getActiveSpreadsheet().toast('Run Setup once via menu.','Setup',10);
		} else {
			menu.addItem('Setup Complete','setupAlreadyDone');
			// Ensure installable onEdit trigger exists post-setup (simple trigger cannot open other spreadsheets)
			ensureInstallableOnEditTrigger();
		}
	}catch(e){
		menu.addItem('Error','setupAlreadyDone');
	}
	menu.addToUi();
	ui.createMenu('📊 Slides').addItem('▶️ Generate Presentation','createPresentationFromSheet').addToUi();
	ui.createMenu('⚙️ Calculations')
		.addItem('Optimize Investor Split','runInvestorSplitOptimization')
		.addSeparator()
		.addItem('Execute Break-Even Analysis','runBreakevenAnalysis')
		.addItem('Reset Break-Even Inputs','resetBreakevenInputs')
		.addToUi();
}
function setupAlreadyDone(){SpreadsheetApp.getActiveSpreadsheet().toast('Setup already completed.','Info',5)}

function runInvestorSplitOptimization(){
	return withGlobalLock('investorSplitOpt',function(){
		var fn='runInvestorSplitOptimization',ui=SpreadsheetApp.getUi(),ss=SpreadsheetApp.getActiveSpreadsheet(),sheet=ss.getSheetByName('Detailed Analysis');
		if(!sheet){ui.alert('Detailed Analysis missing');return}
		try{
			var targetIRRCell=sheet.getRange(TARGET_IRR_INPUT_CELL),val=targetIRRCell.getValue();
			if(typeof val!=='number'||isNaN(val)||val<=0||val>5)throw new Error('Invalid Target IRR in '+TARGET_IRR_INPUT_CELL);
			var finalSplit=calculateInvestorSplitForTargetIRR(sheet,val,INVESTOR_SPLIT_CELL,INVESTOR_IRR_CELL,PROJECT_NET_PROFIT_CELL);
			if(finalSplit!==null){
				ui.alert('Optimization complete. Final Split: '+(finalSplit*100).toFixed(2)+'% Achieved IRR: '+sheet.getRange(INVESTOR_IRR_CELL).getDisplayValue());
			} else {
				ui.alert('Optimization failed.');
			}
		}catch(e){
			ui.alert('Error: '+e.message)
		}
	})
}

function runInitialSetup(){
	var sp=PropertiesService.getScriptProperties(),status=sp.getProperty('SETUP_COMPLETE');
	// Allow re-run if static IDs missing even if status true
	var essentialMissing = ! _looksLikeId(sp.getProperty('DATA_SPREADSHEET_ID')) ||
												 ! _looksLikeId(sp.getProperty('PRELIMINARY_SHEET_ID')) ||
												 ! _looksLikeId(sp.getProperty('SLIDES_TEMPLATE_ID'));
	if(status==='true' && !essentialMissing){
		Logger.log('Setup already complete and essentials present.');
		return;
	}
	var ui=_getUiSafely();
	if(!doesSetupTriggerExist()){ 
		createSetupTrigger(); 
		if(ui) ui.alert('Setup trigger created. Will auto-run shortly.'); else Logger.log('Setup trigger created (no UI context available).');
		return; 
	}
	Logger.log('Running setup (forced='+essentialMissing+')');
	try{
		var ssId=SpreadsheetApp.getActiveSpreadsheet().getId(),file=DriveApp.getFileById(ssId),desc=file.getDescription();
		if(!desc) throw new Error('Spreadsheet description JSON missing');
		var meta=JSON.parse(desc),requestUrl=meta.requestUrl,secret=meta.uniqueSecret; if(!requestUrl||!secret) throw new Error('Missing requestUrl or uniqueSecret');
		var cred=null; function tryFetch(action,payload){ var r=UrlFetchApp.fetch(requestUrl,{method:'post',contentType:'application/json',payload:JSON.stringify(payload),muteHttpExceptions:true}); if(r.getResponseCode()===200){ var body=JSON.parse(r.getContentText()); if(!body.error) return body; throw new Error(body.error||'Unknown credential error'); } throw new Error('HTTP '+r.getResponseCode()); }
		try{ cred=tryFetch('getCredentials',{action:'getCredentials',spreadsheetId:ssId,callbackSecret:secret}); }
		catch(e){ cred=tryFetch('refreshCredentials',{action:'refreshCredentials',spreadsheetId:ssId,originalSecret:secret}); }
		// Log keys present
		Object.keys(cred).forEach(function(k){ if(k.length<40) Logger.log('cred.'+k+' present'); });
		['privateKey','apiKey','staticMapsApiKey','userEmail','gcpProjectId','serviceAccountEmail'].forEach(function(k){ if(cred[k]) sp.setProperty(k,cred[k]); });
		try{
			_storeIdIfBetter(sp,'PRELIMINARY_SHEET_ID',cred.preliminarySheetId);
			_storeIdIfBetter(sp,'DATA_SPREADSHEET_ID',cred.dataSpreadsheetId);
			_storeIdIfBetter(sp,'SLIDES_TEMPLATE_ID',cred.slidesTemplateId);
		}catch(idErr){ Logger.log('Static ID store warn '+idErr.message); }
		var missing=[];
		if(!_looksLikeId(sp.getProperty('DATA_SPREADSHEET_ID'))) missing.push('DATA_SPREADSHEET_ID');
		if(!_looksLikeId(sp.getProperty('PRELIMINARY_SHEET_ID'))) missing.push('PRELIMINARY_SHEET_ID');
		if(!_looksLikeId(sp.getProperty('SLIDES_TEMPLATE_ID'))) missing.push('SLIDES_TEMPLATE_ID');
		if(!sp.getProperty('apiKey')) missing.push('apiKey');
		if(missing.length){
			sp.setProperty('SETUP_COMPLETE','incomplete');
			if(ui) ui.alert('Setup partially complete. Missing: '+missing.join(', ')+'\nEnsure central service returns these fields or set manually then run Run Initial Setup again.');
			Logger.log('Setup incomplete; missing '+missing.join(','));
		}else{
			sp.setProperty('SETUP_COMPLETE','true');
			if(ui) ui.alert('Setup complete.'); else Logger.log('Setup complete (no UI context).');
			deleteOwnSetupTrigger();
			// Create installable onEdit trigger for full auth access to external data spreadsheet
			ensureInstallableOnEditTrigger();
			main(context);
		}
	}catch(e){
		Logger.log('Setup error '+e.message);
		PropertiesService.getScriptProperties().setProperty('SETUP_COMPLETE','fatal_error');
	}
}

// Safe UI retriever (returns null in time-driven / headless contexts)
function _getUiSafely(){ try { return SpreadsheetApp.getUi(); } catch(e){ return null; } }

// Store a fetched static ID if the existing property is absent OR fails heuristic validity.
function _storeIdIfBetter(sp,key,newVal){
	try{
		if(!newVal) return;
		var cur=sp.getProperty(key);
		if(!cur || !_looksLikeId(cur)){
			sp.setProperty(key,newVal);
			Logger.log('[SETUP] Stored/overwrote '+key+' (prev='+(cur?cur.substring(0,6)+'…':'(unset)')+')');
		} else {
			// Keep existing valid-looking value
			if(DEBUG) Logger.log('[DEBUG] _storeIdIfBetter: kept existing '+key);
		}
	}catch(e){Logger.log('[SETUP] _storeIdIfBetter error '+key+' -> '+e.message);}
}

// Manual repair function: call after central service updated to back-fill missing static IDs without resetting everything.
function repairMissingStaticIds(){
	var sp=PropertiesService.getScriptProperties();
	if(_looksLikeId(sp.getProperty('DATA_SPREADSHEET_ID')) && _looksLikeId(sp.getProperty('PRELIMINARY_SHEET_ID')) && _looksLikeId(sp.getProperty('SLIDES_TEMPLATE_ID'))){
		Logger.log('All static IDs already present; nothing to repair.');
		return;
	}
	var ssId=SpreadsheetApp.getActiveSpreadsheet().getId(),file=DriveApp.getFileById(ssId),desc=file.getDescription();
	if(!desc){ Logger.log('repairMissingStaticIds: description missing'); return; }
	try{
		var meta=JSON.parse(desc),requestUrl=meta.requestUrl,secret=meta.uniqueSecret; if(!requestUrl||!secret){ Logger.log('repairMissingStaticIds: missing requestUrl/secret'); return; }
		var payload={action:'getCredentials',spreadsheetId:ssId,callbackSecret:secret};
		var r=UrlFetchApp.fetch(requestUrl,{method:'post',contentType:'application/json',payload:JSON.stringify(payload),muteHttpExceptions:true});
		if(r.getResponseCode()!==200){ Logger.log('repairMissingStaticIds: HTTP '+r.getResponseCode()); return; }
		var body=JSON.parse(r.getContentText());
		if(body.error){ Logger.log('repairMissingStaticIds: '+body.error); return; }
		if(body.preliminarySheetId&&!sp.getProperty('PRELIMINARY_SHEET_ID')) sp.setProperty('PRELIMINARY_SHEET_ID',body.preliminarySheetId);
		_storeIdIfBetter(sp,'DATA_SPREADSHEET_ID',body.dataSpreadsheetId);
		_storeIdIfBetter(sp,'SLIDES_TEMPLATE_ID',body.slidesTemplateId);
		Logger.log('repairMissingStaticIds: updated any missing IDs.');
	}catch(e){ Logger.log('repairMissingStaticIds error '+e.message); }
}
function doesSetupTriggerExist(){try{var t=ScriptApp.getProjectTriggers();for(var i=0;i<t.length;i++)if(t[i].getHandlerFunction()==='runInitialSetup'&&t[i].getEventType()===ScriptApp.EventType.CLOCK)return true;}catch(e){}return false}
function createSetupTrigger(){deleteOwnSetupTrigger();ScriptApp.newTrigger('runInitialSetup').timeBased().after(60000).create()}
function deleteOwnSetupTrigger(){try{ScriptApp.getProjectTriggers().forEach(function(tr){if(tr.getHandlerFunction()==='runInitialSetup'&&tr.getEventType()===ScriptApp.EventType.CLOCK)ScriptApp.deleteTrigger(tr)})}catch(e){}}

function handleSheetEdit(e){var sp=PropertiesService.getScriptProperties();if(sp.getProperty('SETUP_COMPLETE')!=='true')return;var r=e.range,a1=r.getA1Notation(),sheet=r.getSheet();if(sheet.getName()!==context.config.MASTER_SHEET_NAME)return;if(a1===context.config.COMP_RADIUS_CELL)main(context);else if([context.config.DATE_FILTER_CELL,context.config.AGE_FILTER_CELL,context.config.SIZE_FILTER_CELL,context.config.SD_MULTIPLIER_CELL].indexOf(a1)>-1)refilterAndAnalyze(context)}

// Simple trigger wrapper so we don't rely on an installable trigger for basic responsiveness.
function onEdit(e){
	try{
		// Guard: only proceed if event + range present
		if(!e || !e.range){return;}
		var sp=PropertiesService.getScriptProperties();
		// Fast flag (set when we create the installable trigger) so we don't call ScriptApp.getProjectTriggers()
		var installableReady = sp.getProperty('INSTALLABLE_ONEDIT_READY')==='true';
		if(installableReady){
			Logger.log('[onEdit] Installable trigger flag present; skipping (installable will handle logic).');
			return;
		}
		// Fallback (first run before flag set): attempt lightweight detection; may throw in simple trigger (caught -> false)
		if(hasInstallableOnEditTrigger()){
			Logger.log('[onEdit] Detected installable trigger (fallback); skipping.');
			// Set flag so future simple executions exit even faster
			sp.setProperty('INSTALLABLE_ONEDIT_READY','true');
			return;
		}
		Logger.log('[onEdit] (simple) Change detected at '+e.range.getSheet().getName()+'!'+e.range.getA1Notation());
		// Simple triggers have LIMITED auth: avoid any external spreadsheet open attempts or UrlFetch to prevent failures/timeouts.
		// Only allow lightweight filter reapplication when the edit is on filter cells AND data spreadsheet already loaded in context (rare on first run).
		try{
			if(sp.getProperty('SETUP_COMPLETE')==='true'){
				var a1=e.range.getA1Notation();
				if([context.config.DATE_FILTER_CELL,context.config.AGE_FILTER_CELL,context.config.SIZE_FILTER_CELL,context.config.SD_MULTIPLIER_CELL].indexOf(a1)>-1){
					// Re-run local (already imported) filtering only; do NOT call main() here.
					if(context.sheet){
						applyAllFilters(context);
						clearChartDataForHiddenRows(context);
						updateAnalysisOutputs(context);
					}
				}
			}
		}catch(inner){Logger.log('[onEdit] lightweight filtering error '+inner);}
		Logger.log('[onEdit] (simple) Waiting for installable trigger for full processing.');
	}catch(err){
		Logger.log('[onEdit] Error '+err);
	}
}

// Installable onEdit handler (full auth) created post-setup
function onEditInstallable(e){
	try{
		if(!e || !e.range){return;}
		Logger.log('[onEditInstallable] Change detected at '+e.range.getSheet().getName()+'!'+e.range.getA1Notation());
		handleSheetEdit(e);
	}catch(err){
		Logger.log('[onEditInstallable] Error '+err);
	}
}

function ensureInstallableOnEditTrigger(){
	try{
		var triggers=ScriptApp.getProjectTriggers();
		for(var i=0;i<triggers.length;i++) if(triggers[i].getHandlerFunction()==='onEditInstallable'){
			// Ensure flag set if trigger already there (e.g., after code update)
			PropertiesService.getScriptProperties().setProperty('INSTALLABLE_ONEDIT_READY','true');
			return; // already exists
		}
		ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet()).onEdit().create();
		PropertiesService.getScriptProperties().setProperty('INSTALLABLE_ONEDIT_READY','true');
		Logger.log('[TRIGGER] Created installable onEdit trigger & set flag.');
	}catch(e){Logger.log('[TRIGGER] ensureInstallableOnEditTrigger error '+e);}
}

function hasInstallableOnEditTrigger(){
	try{var triggers=ScriptApp.getProjectTriggers();for(var i=0;i<triggers.length;i++) if(triggers[i].getHandlerFunction()==='onEditInstallable') return true;}catch(e){}
	return false;
}

function main(ctx){
	return withGlobalLock('main',function(){
		initializeContext();
		if(!context.sheet) return;
		var sheet=context.sheet,conf=context.config,address=sheet.getRange(conf.ADDRESS_CELL).getValue();
		debugLog('main: address cell '+conf.ADDRESS_CELL+' value="'+address+'"');
		if(!address){ Logger.log('[MAIN] Abort: address cell '+conf.ADDRESS_CELL+' empty.'); return; }
		var radius=sheet.getRange(conf.COMP_RADIUS_CELL).getValue();
		debugLog('main: radius='+radius);
		if(isNaN(radius)||radius<=0){ Logger.log('[MAIN] Abort: radius invalid '+radius+' in '+conf.COMP_RADIUS_CELL); return; }
		var coords=getCoordinatesFromAddress(address);
		if(coords) debugLog('geocode: lat='+coords.lat+' lng='+coords.lng); else debugLog('geocode: failed for address');
		if(!coords){ Logger.log('[MAIN] Abort: geocode failed (likely missing apiKey or SETUP_COMPLETE not true).'); return; }
		var comps=searchComps(coords,radius);
		// Cache subject coords for later analytics (avoid re-geocoding in updateAnalysisOutputs)
		context.subjectCoords=coords;
		debugLog('comps: raw count='+comps.length);
		importCompData(comps,context);
		if(comps.length){
			applyAllFilters(context);
			// Advisory if filters eliminated everything
			var visibleAfter=0;try{for(var r=conf.COMP_RESULTS_START_ROW;r<=sheet.getLastRow();r++) if(!sheet.isRowHiddenByUser(r)) visibleAfter++;}catch(_v){}
			if(visibleAfter===0){
				Logger.log('[MAIN] All comps filtered out. Adjust filters: P11 (SD multiplier) larger, clear P6 (date), lower P8 (year), or increase P9 (size %).');
			}
			clearChartDataForHiddenRows(context);
			updateAnalysisOutputs(context);
			try{var total=sheet.getLastRow()-conf.COMP_RESULTS_START_ROW+1; if(total<0) total=0; var visible=0; for(var r=conf.COMP_RESULTS_START_ROW;r<=sheet.getLastRow();r++) if(!sheet.isRowHiddenByUser(r)) visible++; debugLog('post-filters: visible='+visible+' of '+total);}catch(_cnt){}
		} else {
			Logger.log('[MAIN] No comps found within radius '+radius+'.');
			var es=context.spreadsheet.getSheetByName('Executive Summary');
			if(es){
				es.getRange('G21:K34').clearContent();
				es.getRange('B23:F32').clearContent();
			}
		}
	});
}

// Diagnostics: run manually to see each prerequisite for main()
function diagMainPreconditions(){
	Logger.log('--- diagMainPreconditions ---');
	var sp=PropertiesService.getScriptProperties();
	Logger.log('SETUP_COMPLETE='+(sp.getProperty('SETUP_COMPLETE')||'(unset)'));
	Logger.log('DATA_SPREADSHEET_ID='+(sp.getProperty('DATA_SPREADSHEET_ID')||'(unset)'));
	Logger.log('apiKey present='+(!!sp.getProperty('apiKey')));
	try{initializeContext();}catch(e){Logger.log('initializeContext threw '+e);} 
	Logger.log('context.sheet='+(context.sheet?context.sheet.getName():'(null)'));
	var conf=context.config;
	if(context.sheet){
		var address=context.sheet.getRange(conf.ADDRESS_CELL).getValue();
		Logger.log('Address cell '+conf.ADDRESS_CELL+'="'+address+'"');
		var radius=context.sheet.getRange(conf.COMP_RADIUS_CELL).getValue();
		Logger.log('Radius cell '+conf.COMP_RADIUS_CELL+'='+radius+' (number='+(typeof radius==='number')+')');
	}
	if(context.dataSheet){
		Logger.log('Data sheet rows='+context.dataSheet.getLastRow()+' cols='+context.dataSheet.getLastColumn());
	}else{
		Logger.log('Data sheet not available yet.');
	}
	if(sp.getProperty('SETUP_COMPLETE')==='true'){
		try{var addr=context.sheet?context.sheet.getRange(conf.ADDRESS_CELL).getValue():''; if(addr){var c=getCoordinatesFromAddress(addr); Logger.log('Test geocode result='+(c?c.lat+','+c.lng:'(null)'));}}catch(g){Logger.log('Geocode test error '+g);}
	}else{
		Logger.log('Skipping geocode test because SETUP_COMPLETE != true');
		if(sp.getProperty('SETUP_COMPLETE')==='fatal_error') Logger.log('Hint: run recoverSetupIfComplete() if all required IDs + apiKey are now present.');
	}
	Logger.log('--- end diagMainPreconditions ---');
}

// Recovery helper: if earlier setup marked fatal_error but all essentials now exist, flip to true without rerunning remote fetch.
function recoverSetupIfComplete(){
	var sp=PropertiesService.getScriptProperties();
	var status=sp.getProperty('SETUP_COMPLETE');
	if(status!=='fatal_error'){ Logger.log('recoverSetupIfComplete: status='+status+' (no change).'); return; }
	var missing=[];
	['DATA_SPREADSHEET_ID','PRELIMINARY_SHEET_ID','SLIDES_TEMPLATE_ID'].forEach(function(k){ if(!_looksLikeId(sp.getProperty(k))) missing.push(k); });
	if(!sp.getProperty('apiKey')) missing.push('apiKey');
	if(missing.length){ Logger.log('recoverSetupIfComplete: still missing '+missing.join(', ')); return; }
	sp.setProperty('SETUP_COMPLETE','true');
	Logger.log('recoverSetupIfComplete: SETUP_COMPLETE reset to true. Run main() or edit radius to load comps.');
}

function refilterAndAnalyze(){return withGlobalLock('refilter',function(){initializeContext();if(!context.sheet)return;applyAllFilters(context);SpreadsheetApp.flush();var conf=context.config,sheet=context.sheet,start=conf.COMP_RESULTS_START_ROW,last=sheet.getLastRow(),vis=[];for(var r=start;r<=last;r++)if(!sheet.isRowHiddenByUser(r))vis.push(r);applyFormulasToRows(sheet,vis,32);clearChartDataForHiddenRows(context);updateAnalysisOutputs(context)})}

function applyAllFilters(){
	var sheet=context.sheet,conf=context.config; if(!sheet) return;
	var start=conf.COMP_RESULTS_START_ROW,last=sheet.getLastRow();
	var rowCount=last-start+1; if(rowCount<=0) return;
	// Always unhide first to get a clean baseline
	sheet.showRows(start,rowCount);
	var tAll=Date.now();
	// Pre-fetch user inputs
	var sdMultiplier=sheet.getRange(conf.SD_MULTIPLIER_CELL).getValue();
	var dateThreshold=sheet.getRange('P6').getValue(); // Date filter cell (display P6, internal conf.DATE_FILTER_CELL is P5?) legacy used P6
	var yearThreshold=sheet.getRange('P8').getValue();
	var sizePctRaw=sheet.getRange(conf.SIZE_FILTER_CELL).getValue();
	var subjectSize=sheet.getRange(conf.SUBJECT_SIZE_CELL).getValue();
	if(sizePctRaw>1) sizePctRaw/=100;
	var lowSize=(subjectSize>0)?subjectSize*(1-sizePctRaw):NaN;
	var highSize=(subjectSize>0)?subjectSize*(1+sizePctRaw):NaN;
	if(!isNaN(lowSize)&&!isNaN(highSize)) sheet.getRange(conf.ANNUNCIATOR_CELL).setValue(Math.round(lowSize)+' sqft - '+Math.round(highSize)+' sqft');
	// Pre-fetch needed columns in as few calls as possible
	var sizeCol = sheet.getRange('G'+start+':G'+last).getValues();      // home size
	var yearCol = sheet.getRange('I'+start+':I'+last).getValues();      // year built
	var dateCol = sheet.getRange('J'+start+':J'+last).getValues();      // sale date
	var sdCol   = sheet.getRange(start,14,rowCount,1).getValues();      // metric used for SD filter (col 14)
	// Compute SD stats first (based on all rows) to mirror previous behavior
	var t=Date.now(), sdHidden=0, sdMean=0, sdSD=0, lo=null, hi=null;
	if(!isNaN(sdMultiplier) && sdMultiplier>0){
		var sdVals=[]; for(var i=0;i<rowCount;i++){var v=sdCol[i][0]; if(v && !isNaN(v) && v>0) sdVals.push(Number(v));}
		if(sdVals.length>=2){
			var sum=0; for(var k=0;k<sdVals.length;k++) sum+=sdVals[k];
			sdMean=sum/sdVals.length; var varAcc=0; for(k=0;k<sdVals.length;k++){varAcc+=Math.pow(sdVals[k]-sdMean,2);} var variance=varAcc/sdVals.length; sdSD=Math.sqrt(variance);
			lo=sdMean - sdMultiplier*sdSD; hi=sdMean + sdMultiplier*sdSD;
		}
	}
	var hide=new Array(rowCount); for(var h=0;h<rowCount;h++) hide[h]=false;
	if(lo!=null){
		for(i=0;i<rowCount;i++){var vv=sdCol[i][0]; if( (isNaN(vv)) || vv<lo || vv>hi){ hide[i]=true; sdHidden++; }}
	}
	debugLog('filters(opt): SD pass mean='+(sdMean||0).toFixed?sdMean.toFixed(2):sdMean+' sd='+(sdSD||0).toFixed?sdSD.toFixed(2):sdSD+' lo='+(lo!=null?lo.toFixed(2):'n/a')+' hi='+(hi!=null?hi.toFixed(2):'n/a')+' hid='+sdHidden+' ms='+(Date.now()-t));
	// Date filter
	t=Date.now(); var dateHidden=0; if(dateThreshold instanceof Date){
		for(i=0;i<rowCount;i++){ if(hide[i]) continue; var dv=dateCol[i][0]; if(dv instanceof Date && dv<dateThreshold){ hide[i]=true; dateHidden++; }}
	}
	debugLog('filters(opt): Date hid='+dateHidden+' ms='+(Date.now()-t));
	// Age filter
	t=Date.now(); var ageHidden=0; if(!isNaN(yearThreshold) && yearThreshold>=1900){
		for(i=0;i<rowCount;i++){ if(hide[i]) continue; var yv=yearCol[i][0]; if(isNaN(yv) || yv<yearThreshold){ hide[i]=true; ageHidden++; }}
	}
	debugLog('filters(opt): Age hid='+ageHidden+' ms='+(Date.now()-t));
	// Size filter
	t=Date.now(); var sizeHidden=0; if(subjectSize>0 && !isNaN(lowSize) && !isNaN(highSize)){
		for(i=0;i<rowCount;i++){ if(hide[i]) continue; var sv=sizeCol[i][0]; if(isNaN(sv) || sv<lowSize || sv>highSize){ hide[i]=true; sizeHidden++; }}
	}
	debugLog('filters(opt): Size hid='+sizeHidden+' range='+Math.round(lowSize)+'-'+Math.round(highSize)+' ms='+(Date.now()-t));
	// Apply hides in contiguous batches
	t=Date.now();
	var batchStart=null,batchLen=0,totalHidden=0;
	for(i=0;i<rowCount;i++){
		if(hide[i]){ totalHidden++; if(batchStart===null){ batchStart=start+i; batchLen=1; } else { batchLen++; } }
		else if(batchStart!==null){ sheet.hideRows(batchStart,batchLen); batchStart=null; batchLen=0; }
	}
	if(batchStart!==null){ sheet.hideRows(batchStart,batchLen); }
	var visible=rowCount-totalHidden;
	debugLog('filters(opt): applied hidden='+totalHidden+' visible='+visible+' hideOps(ms)='+(Date.now()-t)+' totalMs='+(Date.now()-tAll));
}
function applySDFilter(ctx,m){var sheet=ctx.sheet,conf=ctx.config,start=conf.COMP_RESULTS_START_ROW,end=sheet.getLastRow(),col=14;if(end<start)return 0;var vals=sheet.getRange(start,col,end-start+1,1).getValues(),arr=[];for(var i=0;i<vals.length;i++){var row=start+i;if(!sheet.isRowHiddenByUser(row)){var v=vals[i][0];if(v&&!isNaN(v)&&v>0)arr.push(Number(v))}}if(arr.length<2)return 0;var mean=arr.reduce((a,b)=>a+b,0)/arr.length,varc=arr.reduce((s,v)=>s+Math.pow(v-mean,2),0)/arr.length,sd=Math.sqrt(varc),lo=mean-m*sd,hi=mean+m*sd,hidden=0;for(i=0;i<vals.length;i++){var row2=start+i;if(!sheet.isRowHiddenByUser(row2)){var v2=vals[i][0];if(isNaN(v2)||v2<lo||v2>hi){sheet.hideRows(row2);hidden++;}}}return hidden}
function applyDateFilter(ctx){var sheet=ctx.sheet,conf=ctx.config,d=sheet.getRange('P6').getValue();if(!(d instanceof Date))return 0;var last=sheet.getLastRow(),rng=sheet.getRange('J'+conf.COMP_RESULTS_START_ROW+':J'+last).getValues(),hidden=0;for(var i=0;i<rng.length;i++){var sd=rng[i][0],row=conf.COMP_RESULTS_START_ROW+i;if(sd instanceof Date && sd<d && !sheet.isRowHiddenByUser(row)){sheet.hideRows(row);hidden++;}}return hidden}
function applyAgeFilter(ctx){var sheet=ctx.sheet,conf=ctx.config,y=sheet.getRange('P8').getValue();if(isNaN(y)||y<1900)return 0;var last=sheet.getLastRow(),rng=sheet.getRange('I'+conf.COMP_RESULTS_START_ROW+':I'+last).getValues(),hidden=0;for(var i=0;i<rng.length;i++){var v=rng[i][0],row=conf.COMP_RESULTS_START_ROW+i;if((isNaN(v)||v<y) && !sheet.isRowHiddenByUser(row)){sheet.hideRows(row);hidden++;}}return hidden}
function applySizeFilter(ctx){var sheet=ctx.sheet,conf=ctx.config,size=sheet.getRange(conf.SUBJECT_SIZE_CELL).getValue();if(isNaN(size)||size<=0)return 0;var pct=sheet.getRange(conf.SIZE_FILTER_CELL).getValue();if(pct>1)pct/=100;var low=size*(1-pct),high=size*(1+pct);sheet.getRange(conf.ANNUNCIATOR_CELL).setValue(Math.round(low)+' sqft - '+Math.round(high)+' sqft');var last=sheet.getLastRow(),rng=sheet.getRange('G'+conf.COMP_RESULTS_START_ROW+':G'+last).getValues(),hidden=0;for(var i=0;i<rng.length;i++){var v=rng[i][0],row=conf.COMP_RESULTS_START_ROW+i;if((isNaN(v)||v<low||v>high) && !sheet.isRowHiddenByUser(row)){sheet.hideRows(row);hidden++;}}return hidden}

// Diagnostic: inspect first row headers & detect coordinate column used (row[41])
function diagCompHeaders(){
	try{
		initializeContext();
		if(!context.dataSheet){Logger.log('diagCompHeaders: dataSheet not available');return;}
		var header=context.dataSheet.getRange(1,1,1,60).getValues()[0];
		for(var i=0;i<header.length;i++){
			var h=header[i];
			if(h!==''&&h!=null) Logger.log('HDR['+i+']='+(h.length>40?h.substring(0,40)+'…':h));
		}
		// Sample a few rows for the supposed coordinate column 41
		for(var r=2;r<=6;r++){
			var val=context.dataSheet.getRange(r,42).getValue(); // 42 is 1-based column index for row[41]
			if(val) Logger.log('Row '+r+' col42 value='+val);
		}
	}catch(e){Logger.log('diagCompHeaders error '+e.message);}
}

function getCoordinatesFromAddress(address){var sp=PropertiesService.getScriptProperties();if(sp.getProperty('SETUP_COMPLETE')!=='true')return null;var key=sp.getProperty('apiKey');if(!key)return null;var url='https://maps.googleapis.com/maps/api/geocode/json?address='+encodeURIComponent(address)+'&key='+key;try{var r=UrlFetchApp.fetch(url);var d=JSON.parse(r.getContentText());if(r.getResponseCode()===200&&d.status==='OK'){var loc=d.results[0].geometry.location;return{lat:loc.lat,lng:loc.lng}}}catch(e){Logger.log('geocode err '+e.message)}return null}
function searchComps(coords,radius){if(!context.dataSpreadsheet)context.dataSpreadsheet=openDataSpreadsheet();if(!context.dataSheet)context.dataSheet=getDataSheet(context.dataSpreadsheet);if(!context.dataSheet)return[];var values=context.dataSheet.getDataRange().getValues();return findCompsWithinRadius(values,coords,radius)}
function openDataSpreadsheet(){try{return SpreadsheetApp.openById(context.config.DATA_SPREADSHEET_ID)}catch(e){return null}}
function getDataSheet(ss){return ss?ss.getSheetByName(context.config.DATA_SHEET_NAME):null}
function findCompsWithinRadius(data,coords,radius){var out=[];for(var i=1;i<data.length;i++){var row=data[i];if(!row||row.length<=41)continue;var ll=row[41];if(!ll)continue;var parts=ll.split(','),lat=parseFloat(parts[0]),lng=parseFloat(parts[1]);if(isNaN(lat)||isNaN(lng))continue;var dist=calculateDistance(coords.lat,coords.lng,lat,lng);if(dist<=radius)out.push({address:row[0],city:row[2],state:row[3],zip:row[4],beds:row[21],baths:row[22],buildingSqft:row[23],lotSize:row[24],yearBuilt:row[25],date:row[27],price:row[36],distance:dist,lat:lat,lng:lng})}return out}
function calculateDistance(a,b,c,d){var toR=x=>x*Math.PI/180,R=3958.8,dLat=toR(c-a),dLng=toR(d-b),A=Math.sin(dLat/2)**2+Math.cos(toR(a))*Math.cos(toR(c))*Math.sin(dLng/2)**2;return 2*Math.atan2(Math.sqrt(A),Math.sqrt(1-A))*R}

function importCompData(props,ctx){
  var sheet=ctx.sheet,conf=ctx.config,start=conf.COMP_RESULTS_START_ROW;
  if(sheet.getLastRow()>=start)sheet.getRange(start,1,Math.max(1,sheet.getLastRow()-start+1),23).clearContent();
  if(!props.length){ctx.lastImportCount=0;return;}
  var rows=props.map(p=>[p.address,p.city,p.state,p.zip,p.beds,p.baths,p.buildingSqft,p.lotSize,p.yearBuilt,p.date,p.price,'',p.distance!=null?Number(p.distance).toFixed(2):'', '',p.lat,p.lng]);
  var rng=sheet.getRange(start,1,rows.length,16);rng.setValues(rows);ctx.lastImportCount=rows.length;
  try{var templateRow=32,cols=[14,18,19,20,21,22,23];cols.forEach(function(c){var f=sheet.getRange(templateRow,c).getFormulaR1C1();if(f)sheet.getRange(start,c,rows.length,1).setFormulaR1C1(f)})}catch(e){}
  try{sheet.getRange(start,7,rows.length,2).setNumberFormat('#,##0');sheet.getRange(start,11,rows.length,1).setNumberFormat('$#,##0');sheet.getRange(start,13,rows.length,1).setNumberFormat('0.00');sheet.getRange(start,15,rows.length,2).setNumberFormat('0.000000')}catch(e){}
}

function updateAnalysisOutputs(ctx){
	try{
		var tAll=Date.now();
		var conf=ctx.config, sheet=ctx.sheet;
		if(!sheet) return;
		// Build / reuse visible comps cache (bulk read)
		var cache=collectVisibleCompData(ctx);
		if(cache.rows.length===0){ debugLog('updateAnalysisOutputs: skipped (no visible comps).'); return; }
		var t1=Date.now();
		calculateTrendlineAndPricePerSqft(ctx,cache); // now supports cached dataset
		var t2=Date.now();
		updateChartWithMargins(ctx,cache);
		var t3=Date.now();
		try{enforceChartStyles(ctx);}catch(_s){debugLog('enforceChartStyles error '+_s);} // style stabilization
		populateExecutiveSummaryCompsTable(ctx,cache);
		var t4=Date.now();
		// Map generation (reuse cached coords + subjectCoords when present)
		var sp=PropertiesService.getScriptProperties(),mapKey=sp.getProperty('staticMapsApiKey');
		if(mapKey){
			var coords=ctx.subjectCoords; // prefer cached
			if(!coords){
				var subj=sheet.getRange(conf.ADDRESS_CELL).getValue();
				coords=getCoordinatesFromAddress(subj);
				ctx.subjectCoords=coords; // cache for future
			}
			if(coords){
				var markers=cache.latLngs;
				var mapUrl=generateStaticMapUrl(coords,markers,mapKey);
				var exec=ctx.spreadsheet.getSheetByName('Executive Summary');
				if(exec&&mapUrl) insertMapIntoSheet(exec,mapUrl,'G21');
				debugLog('map: markers='+markers.length+' urlLen='+(mapUrl?mapUrl.length:0));
			}
		}
		var t5=Date.now();
		updatePreliminarySheet();
		debugLog('updateAnalysisOutputs complete timings(ms): cache='+(t1-tAll)+' trend='+(t2-t1)+' chart='+(t3-t2)+' summary='+(t4-t3)+' map='+(t5-t4)+' total='+(Date.now()-tAll));
	}catch(e){
		Logger.log('update outputs error '+e.message);
	}
}

// Collect visible comp row data with a single bulk range read; caches on context until next filter/import.
function collectVisibleCompData(ctx){
	var conf=ctx.config, sheet=ctx.sheet, start=conf.COMP_RESULTS_START_ROW, last=sheet.getLastRow();
	var rowCount=last-start+1; if(rowCount<=0) return {rows:[],sizes:[],pps:[],prices:[],latLngs:[]};
	// If cache exists and last row unchanged, reuse
	if(ctx._visibleCompCache && ctx._visibleCompCache._stamp===last) return ctx._visibleCompCache;
	var values=sheet.getRange(start,1,rowCount,16).getValues(); // columns A..P (ensure lat/lng included)
	var rows=[],sizes=[],pps=[],prices=[],latLngs=[];
	for(var i=0;i<rowCount;i++){
		var rowNum=start+i;
		if(sheet.isRowHiddenByUser(rowNum)) continue; // unavoidable per-row check (no bulk API)
		var row=values[i];
		var address=row[0];
		var size=row[6];
		var salePrice=row[10];
		var distance=row[12];
		var pricePerSqft=row[13]; // col 14 (N)
		var lat=row[14]; // O
		var lng=row[15]; // P
		rows.push({row:rowNum,address:address,size:size,price:salePrice,distance:distance,pricePerSqft:pricePerSqft,lat:lat,lng:lng});
		if(size>0 && pricePerSqft>0){ sizes.push(size); pps.push(pricePerSqft); }
		if(size>0 && salePrice>0){ prices.push(salePrice); }
		if(lat && lng) latLngs.push({lat:lat,lng:lng});
	}
	var cache={rows:rows,sizes:sizes,pps:pps,prices:prices,latLngs:latLngs,_stamp:last};
	ctx._visibleCompCache=cache;
	return cache;
}

function calculateTrendlineAndPricePerSqft(ctx,cache){
	var sheet=ctx.sheet;
	cache=cache||collectVisibleCompData(ctx);
	var sizes=cache.sizes, pps=cache.pps;
	if(!sizes.length) return;
	var N=sizes.length,sumX=sizes.reduce((a,b)=>a+b,0),sumY=pps.reduce((a,b)=>a+b,0),sumXY=sizes.reduce((s,x,i)=>s+x*pps[i],0),sumX2=sizes.reduce((s,x)=>s+x*x,0),den=(N*sumX2 - sumX*sumX); if(den===0) return; var slope=(N*sumXY - sumX*sumY)/den,inter=(sumY - slope*sumX)/N;
	var subject=sheet.getRange('G3:G5').getValues(),out=[];subject.forEach(function(r){var hs=r[0];out.push([hs>0?((slope*hs)+inter).toFixed(2):''])});
	sheet.getRange('N3:N5').setValues(out);
	debugLog('trendline: points='+N+' slope='+slope.toFixed(6)+' intercept='+inter.toFixed(2));
}

function updateChartWithMargins(ctx,cache){
	var sheet=ctx.sheet,charts=sheet.getCharts(),hp=null,ppsChart=null;
	charts.forEach(function(c){
		var t=c.getOptions().get('title');
		if(t==='Home Price ($) x Home Size (sq.ft.)') hp=c; else if(t==='Home Prices per Square Foot ($) x Home Size (sq.ft.)') ppsChart=c;
	});
	if(!hp||!ppsChart) return;
	cache=cache||collectVisibleCompData(ctx);
	if(cache.rows.length===0) return;
	var sizes=[], ppsArr=[], prices=[];
	for(var i=0;i<cache.rows.length;i++){
		var r=cache.rows[i];
		if(r.size>0 && r.pricePerSqft>0 && r.price>0){
			sizes.push(r.size); ppsArr.push(r.pricePerSqft); prices.push(r.price);
		}
	}
	if(!sizes.length) return;
	var minS=Math.min.apply(null,sizes),maxS=Math.max.apply(null,sizes),minPPS=Math.min.apply(null,ppsArr),maxPPS=Math.max.apply(null,ppsArr),minPr=Math.min.apply(null,prices),maxPr=Math.max.apply(null,prices),mS=300,mP=300,mPr=300000;
	sheet.updateChart(hp.modify()
		.setOption('hAxis.minValue',Math.max(0,minS-mS))
		.setOption('hAxis.maxValue',maxS+mS)
		.setOption('vAxis.minValue',Math.max(0,minPr-mPr))
		.setOption('vAxis.maxValue',maxPr+mPr)
		.build());
	sheet.updateChart(ppsChart.modify()
		.setOption('hAxis.minValue',Math.max(0,minS-mS))
		.setOption('hAxis.maxValue',maxS+mS)
		.setOption('vAxis.minValue',Math.max(0,minPPS-mP))
		.setOption('vAxis.maxValue',maxPPS+mP)
		.build());
}

// Enforce consistent series colors / point sizes for the two scatter charts.
function enforceChartStyles(ctx){
	var sheet=ctx.sheet;if(!sheet)return;var charts=sheet.getCharts();
	charts.forEach(function(c){
		var title=c.getOptions().get('title');
		if(title==='Home Price ($) x Home Size (sq.ft.)' || title==='Home Prices per Square Foot ($) x Home Size (sq.ft.)'){
			var builder=c.modify();
			// Apply legend position preference
			if(CHART_LEGEND_POSITION) builder.setOption('legend',{position:CHART_LEGEND_POSITION});
			// Iterate defined styles
			for(var idx in CHART_SERIES_STYLE){
				if(!CHART_SERIES_STYLE.hasOwnProperty(idx)) continue; var num=parseInt(idx,10),style=CHART_SERIES_STYLE[idx];
				var opt={}; if(style.color) opt.color=style.color; if(style.pointSize!=null) opt.pointSize=style.pointSize;
				// Apply series-specific options under series.<index>
				Object.keys(opt).forEach(function(k){builder.setOption('series.'+num+'.'+k,opt[k]);});
			}
			// If subject series (0) should show labels (address) ensure data label style optional future enhancement
			sheet.updateChart(builder.build());
		}
	});
}

function populateExecutiveSummaryCompsTable(ctx,cache){
	var exec=ctx.spreadsheet.getSheetByName('Executive Summary'); if(!exec) return;
	cache=cache||collectVisibleCompData(ctx);
	if(cache.rows.length===0){ exec.getRange('B23:F32').clearContent(); return; }
	// Sort by distance (distance may be string -> parseFloat)
	var sorted=cache.rows.slice().sort(function(a,b){ var da=parseFloat(a.distance)||Infinity, db=parseFloat(b.distance)||Infinity; return da-db; });
	var top=sorted.slice(0,10);
	var table=top.map(function(c){ return [c.address||'N/A', c.size||null, c.price||null, c.distance!=null?Number(c.distance).toFixed(2):null, c.pricePerSqft||null]; });
	var rng=exec.getRange('B23:F32'); rng.clearContent(); if(table.length) exec.getRange(23,2,table.length,5).setValues(table);
}

function clearChartDataForHiddenRows(ctx){var sheet=ctx.sheet,start=ctx.config.COMP_RESULTS_START_ROW,last=sheet.getLastRow(),cols=[18,20,23];for(var r=start;r<=last;r++)if(sheet.isRowHiddenByUser(r))cols.forEach(c=>{if(c<=sheet.getMaxColumns())sheet.getRange(r,c).clearContent()})}
function applyFormulasToRows(sheet,rows,templateRow){var cols=[14,18,19,20,21,22,23],formulas={};cols.forEach(c=>{try{formulas[c]=sheet.getRange(templateRow,c).getFormulaR1C1()}catch(e){formulas[c]=null}});rows.forEach(r=>{cols.forEach(c=>{if(formulas[c])sheet.getRange(r,c).setFormulaR1C1(formulas[c])})});SpreadsheetApp.flush()}

function generateStaticMapUrl(subject,compCoords,key,w=640,h=400){
	if(!subject||!key) return null;
	var base='https://maps.googleapis.com/maps/api/staticmap',
		p=['size='+w+'x'+h,'maptype=roadmap'],
		vis=[subject.lat+','+subject.lng];
	(compCoords||[]).forEach(function(c){ if(c.lat&&c.lng) vis.push(c.lat+','+c.lng); });
	if(vis.length>1){
		p.push('visible='+vis.join('%7C'));
	} else {
		p.push('center='+subject.lat+','+subject.lng);
		p.push('zoom=15');
	}
	p.push('markers=color:red%7Clabel:S%7C'+subject.lat+','+subject.lng);
	(compCoords||[]).forEach(function(c,i){ if(c.lat&&c.lng) p.push('markers=color:green%7Clabel:'+(i<9?i+1:'C')+'%7C'+c.lat+','+c.lng); });
	p.push('key='+key);
	return base+'?'+p.join('&');
}
function insertMapIntoSheet(sheet,url,a1,merged){try{if(merged)sheet.getRange(merged).clearContent();sheet.getRange(a1).setFormula('=IMAGE("'+url+'",2)')}catch(e){}}

function normalizeAddress(a){if(!a)return'';return String(a).toLowerCase().replace(/[.,\/#!$%\^&\*;:{}=\-_`~()]/g,'').replace(/\b(street|str)\b/g,'st').replace(/\b(road|rd)\b/g,'rd').replace(/\b(avenue|ave)\b/g,'ave').replace(/\b(boulevard|blvd)\b/g,'blvd').replace(/\b(lane|ln)\b/g,'ln').replace(/\b(place|plz|pl)\b/g,'pl').replace(/\b(court|ct)\b/g,'ct').replace(/\b(cove|cv)\b/g,'cv').replace(/\b(trail|trl)\b/g,'trl').replace(/\b(drive|dr)\b/g,'dr').replace(/\b(circle|cir)\b/g,'cir').replace(/\b(unit|apt|suite)\b/g,'#').replace(/\bnorth\b/g,'n').replace(/\bsouth\b/g,'s').replace(/\beast\b/g,'e').replace(/\bwest\b/g,'w').trim().replace(/\s+/g,' ')}
function findRowByAddressInPreliminary(sheet,address){var tgt=normalizeAddress(address);if(!tgt)return null;var rng=sheet.getRange('B4:B'+sheet.getLastRow()).getValues();for(var i=0;i<rng.length;i++){var cur=rng[i][0];if(cur&&normalizeAddress(cur)===tgt)return i+4}return null}
function updatePreliminarySheet(){
	var ss=SpreadsheetApp.getActiveSpreadsheet(),das=ss.getSheetByName('Detailed Analysis'),exec=ss.getSheetByName('Executive Summary');
	if(!das||!exec)return;var prelimId=getProp('PRELIMINARY_SHEET_ID','');
	if(!_looksLikeId(prelimId))return;var simple=das.getRange(METRIC_CELLS.SIMPLE_ADDRESS).getValue();if(!simple)return;
	var net=exec.getRange(METRIC_CELLS.NET_PROFIT).getValue(),roi=exec.getRange(METRIC_CELLS.ROI).getValue(),margin=exec.getRange(METRIC_CELLS.MARGIN).getValue();
	if(net===''&&roi===''&&margin==='')return;
	var prelim;try{prelim=SpreadsheetApp.openById(prelimId);}catch(e){Logger.log('updatePreliminarySheet: open prelim failed '+e);return;}
	var sheet=prelim.getSheetByName(PRELIMINARY_SHEET_NAME);if(!sheet)return;
	var row=findRowByAddressInPreliminary(sheet,simple);if(row)sheet.getRange(row,23,1,3).setValues([[net,roi,margin]]);
}

function calculateInvestorSplitForTargetIRR(sheet,target,splitCell,irrCell,profitCell){
	var MAX=100,TOL=0.001,step=0.01,profit=sheet.getRange(profitCell).getValue();
	if(!(profit>0)) return null;
	var split=sheet.getRange(splitCell).getValue();
	if(!(split>0&&split<1)) split=.5;
	for(var i=0;i<MAX;i++){
		sheet.getRange(splitCell).setValue(split);
		SpreadsheetApp.flush();
		Utilities.sleep(400);
		var irr=sheet.getRange(irrCell).getValue();
		if(typeof irr!=='number'||isNaN(irr)){
			split+=step*0.1;
			split=Math.min(.99,Math.max(.01,split));
			continue;
		}
		var diff=irr-target;
		if(Math.abs(diff)<=TOL) return split;
		if(diff>0){
			split-=step;
		} else {
			split+=step;
		}
		split=Math.min(.99,Math.max(.01,split));
		if(Math.abs(diff)<.05){
			step=.005;
		} else {
			step=.01;
		}
	}
	return split;
}

// Presentation generation (kept concise)
function createPresentationFromSheet(){
	var ui=SpreadsheetApp.getUi();
	var slidesId=getProp('SLIDES_TEMPLATE_ID','');
	if(!_looksLikeId(slidesId)){
		ui.alert('Slides template ID missing');
		return;
	}
	try{
		var ss=SpreadsheetApp.getActiveSpreadsheet(),
			da=ss.getSheetByName('Detailed Analysis'),
			chartSheet=ss.getSheetByName(CHART_SHEET_NAME);
		if(!da||!chartSheet) throw new Error('Required sheets missing');
		var addr=da.getRange('B6').getValue();
		if(!addr) throw new Error('Simple address missing B6');
		var irr=da.getRange('B153').getDisplayValue(),
			roi=da.getRange('B151').getDisplayValue(),
			multiple=da.getRange('B152').getDisplayValue(),
			net=da.getRange('B138').getDisplayValue(),
			gross=da.getRange('B135').getDisplayValue();
		var file=DriveApp.getFileById(ss.getId()),
			parent=file.getParents().next(),
			copy=DriveApp.getFileById(slidesId).makeCopy(addr+' - Investor Summary',parent),
			pres=SlidesApp.openById(copy.getId()),
			slides=pres.getSlides();
		if(slides.length<4) throw new Error('Template missing slide 4');
		var slide=slides[3];
		slide.replaceAllText('{{TARGET_IRR}}',irr||'N/A');
		slide.replaceAllText('{{TARGET_ROI}}',roi||'N/A');
		slide.replaceAllText('{{TARGET_MULTIPLE}}',multiple||'N/A');
		slide.replaceAllText('{{NET_PROFIT}}',net||'N/A');
		slide.replaceAllText('{{GROSS_PROFIT}}',gross||'N/A');
		var charts=chartSheet.getCharts(),c1=null,c2=null;
		charts.forEach(function(c){
			var t=c.getOptions().get('title');
			if(t===PIE_CHART_1_TITLE){
				c1=c;
			} else if(t===PIE_CHART_2_TITLE){
				c2=c;
			}
		});
		var w=pres.getPageWidth()*0.45,
			h=w*0.75,
			gap=-25,
			slideW=pres.getPageWidth(),
			slideH=pres.getPageHeight(),
			total=2*w+gap,
			left=(slideW-total)/2,
			top=slideH-h-10;
		if(c1){ slide.insertSheetsChartAsImage(c1,left,top,w,h); }
		if(c2){ slide.insertSheetsChartAsImage(c2,left+w+gap,top,w,h); }
		pres.saveAndClose();
		var exec=ss.getSheetByName('Executive Summary');
		if(exec){ exec.getRange('A124').insertHyperlink(copy.getUrl(),copy.getName()); }
		ui.alert('Presentation generated.');
	}catch(e){
		ui.alert('Error: '+e.message);
	}
}

// Map image helpers rely on functions above

// Entry points already declared above
