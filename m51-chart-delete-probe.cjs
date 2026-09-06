'use strict'

function chartDeleteProbeCommand(spec){
 function methods(o){var out=[];if(!o)return out;var seen={};var p=o;for(var depth=0;p&&depth<8;depth++,p=Object.getPrototypeOf(p)){var ns=[];try{ns=Object.getOwnPropertyNames(p)}catch(_){}for(var i=0;i<ns.length;i++){var n=ns[i];if(seen[n])continue;seen[n]=1;try{if(typeof o[n]==='function')out.push(n)}catch(_){}}}return out.sort()}
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function safe(o,n){try{return has(o,n)?o[n]():null}catch(e){return 'ERROR:'+String(e&&e.message?e.message:e)}}
 try{
  var sh=has(Api,'GetSheet')?Api.GetSheet(spec.sheet):null;if(!sh)return {ok:false,error:'sheet unavailable'}
  var charts=has(sh,'GetAllCharts')?(sh.GetAllCharts()||[]):[]
  var drawings=has(sh,'GetAllDrawings')?(sh.GetAllDrawings()||[]):[]
  var c=charts.length?charts[0]:null,d=drawings.length?drawings[0]:null
  var workbook=null;try{workbook=has(Api,'GetWorkbook')?Api.GetWorkbook():null}catch(_){}
  var workbookByName=null,chartName=c?safe(c,'GetName'):null
  if(workbook&&chartName&&has(workbook,'GetDrawingsByName')){try{workbookByName=workbook.GetDrawingsByName(chartName)}catch(e){workbookByName='ERROR:'+String(e&&e.message?e.message:e)}}
  function desc(x){if(!x)return null;return {name:safe(x,'GetName'),classType:safe(x,'GetClassType'),methods:methods(x),hasDelete:has(x,'Delete'),hasRemove:has(x,'Remove')}}
  var byNameDesc=null
  if(Array.isArray(workbookByName))byNameDesc=workbookByName.map(desc);else if(workbookByName&&typeof workbookByName==='object')byNameDesc=desc(workbookByName);else byNameDesc=workbookByName
  return {ok:true,source:'live-coedit-editor',sheet:spec.sheet,counts:{charts:charts.length,drawings:drawings.length},worksheetMethods:methods(sh),apiMethods:methods(Api),chart:desc(c),drawing:desc(d),sameObject:c&&d?c===d:null,workbook:{available:!!workbook,methods:methods(workbook),hasGetDrawingsByName:has(workbook,'GetDrawingsByName'),byName:byNameDesc}}
 }catch(e){return {ok:false,source:'live-coedit-editor',error:String(e&&e.message?e.message:e)}}
}

async function runChartDeleteProbeInFrame(frame,apiHely,sheet,timeoutMs=15000){const body=`return (${chartDeleteProbeCommand.toString()})(${JSON.stringify({sheet})});`;return frame.evaluate(({u,body,timeout})=>new Promise(resolve=>{const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};try{editor.callCommand(new Function(body),false,v=>finish(v))}catch(e){finish({ok:false,error:String(e&&e.message?e.message:e)})}setTimeout(()=>finish({ok:false,error:'callback-timeout'}),timeout)}),{u:apiHely,body,timeout:timeoutMs})}
module.exports={chartDeleteProbeCommand,runChartDeleteProbeInFrame}
