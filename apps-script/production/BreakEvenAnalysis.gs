// Production copy of Break-Even Analysis core functions (trimmed). Merge into geoComp if preferred.
// Expects menu entries already added by geoComp. Uses Scenario Distribution + Detailed Analysis sheets.

function runBreakevenAnalysis(){ if(typeof calculateBreakevenForSingleVariable!=='function'){ Logger.log('Break-even support not fully loaded.'); return; } return _runBreakevenAnalysisImpl(); }

// Below is a compact rendition of the logic; keep originals for readability.
function _runBreakevenAnalysisImpl(){ var fn='runBreakevenAnalysis'; Logger.log('['+fn+'] start'); var ss=SpreadsheetApp.getActiveSpreadsheet(); var scenario=ss.getSheetByName('Scenario Distribution'); var da=ss.getSheetByName('Detailed Analysis'); if(!scenario||!da){ Logger.log('['+fn+'] missing sheets'); return; }
 var baseline={ cost: da.getRange('B82').getValue(), price: da.getRange('B56').getValue(), timeline: da.getRange('B32').getValue() };
 var solveVar=null,label='',targetCell='',resultCell='';
 if(scenario.getRange('I4').isChecked()){ solveVar='cost'; label='Max Construction Cost'; targetCell='B82'; resultCell='K6'; }
 else if(scenario.getRange('M4').isChecked()){ solveVar='price'; label='Min Sales Price'; targetCell='B56'; resultCell='O6'; }
 else if(scenario.getRange('Q4').isChecked()){ solveVar='timeline'; label='Max Timeline'; targetCell='B32'; resultCell='S6'; }
 if(!solveVar){ Logger.log('['+fn+'] nothing selected'); return; }
 var desiredCost=scenario.getRange('J6').getValue(); if(solveVar!=='cost' && desiredCost!==''){ da.getRange('B82').setValue(desiredCost); baseline.cost=desiredCost; }
 var desiredPrice=scenario.getRange('N6').getValue(); if(solveVar!=='price' && desiredPrice!==''){ da.getRange('B56').setValue(desiredPrice); baseline.price=desiredPrice; }
 var desiredTimeline=scenario.getRange('R6').getValue(); if(solveVar!=='timeline' && desiredTimeline!==''){ da.getRange('B32').setValue(desiredTimeline); baseline.timeline=desiredTimeline; }
 SpreadsheetApp.flush(); Utilities.sleep(300);
 var currentVal=da.getRange(targetCell).getValue(); var dir=(solveVar==='price'?-1:1); var step=Math.max( (typeof currentVal==='number'? currentVal*0.01:1), (solveVar==='timeline'?1:100) );
 var targetProfitCell='B137'; var maxIter=200, tol=100; var value=currentVal; var lastDiff=null;
 for(var i=0;i<maxIter;i++){ da.getRange(targetCell).setValue(value); SpreadsheetApp.flush(); Utilities.sleep(400); var profit=da.getRange(targetProfitCell).getValue(); if(typeof profit!=='number'||isNaN(profit)){ Logger.log('bad profit'); break; }
 var diff=profit-0; if(Math.abs(diff)<=tol){ Logger.log('breakeven reached'); break; }
 if(lastDiff!==null && Math.sign(diff)!==Math.sign(lastDiff)){ dir*=-1; step/=2; }
 value += step*dir; if(step < (solveVar==='timeline'?0.5:1)){ break; } lastDiff=diff; }
 da.getRange(resultCell).setValue(value); Logger.log('['+fn+'] done'); }
