'use strict'
function probe(spec){
 function has(o,n){return !!o&&typeof o[n]==='function'}
 function methods(o){var r=[],s={},p=o;for(var d=0;p&&d<8;d++,p=Object.getPrototypeOf(p)){var ns=[];try{ns=Object.getOwnPropertyNames(p)}catch(_){}for(var i=0;i<ns.length;i++){var n=ns[i];if(s[n])continue;s[n]=1;try{if(typeof o[n]==='function')r.push(n)}catch(_){}}}return r.sort()}
 function safe(o,n){try{return has(o,n)?o[n]():null}catch(e){return 'ERROR:'+String(e&&e.message?e.message:e)}}
 function desc(o){return o?{name:safe(o,'GetName'),classType:safe(o,'GetClassType'),methods:methods(o),hasDelete:has(o,'Delete'),hasRemove:has(o,'Remove'),hasSelect:has(o,'Select'),hasUnselect:has(o,'Unselect')}:null}
 try{
  var sh=Api.GetSheet(spec.sheet);if(!sh)return {ok:false,error:'sheet unavailable'}
  var before=sh.GetAllCharts()||[],name='M51_PROBE_'+Date.now(),chart=sh.AddChart("'"+String(spec.sheet).replace(/'/g,"''")+"'!"+spec.range,false,'bar',1,3600000,2200000,5,0,0,0);if(!chart)return {ok:false,error:'AddChart returned null'}
  if(has(chart,'SetName'))chart.SetName(name)
  var drawings=has(sh,'GetAllDrawings')?(sh.GetAllDrawings()||[]):[],match=null;for(var i=0;i<drawings.length;i++)if(safe(drawings[i],'GetName')===name){match=drawings[i];break}
  var wb=null;try{wb=has(Api,'GetActiveWorkbook')?Api.GetActiveWorkbook():null}catch(_){}
  var byName=null;if(wb&&has(wb,'GetDrawingsByName')){try{byName=wb.GetDrawingsByName([name])||[]}catch(e){byName='ERROR:'+String(e&&e.message?e.message:e)}}
  var byDesc=Array.isArray(byName)?byName.map(desc):byName
  return {ok:true,source:'live-coedit-editor',name:name,beforeCount:before.length,afterCount:(sh.GetAllCharts()||[]).length,chart:desc(chart),drawing:desc(match),sameObject:!!match&&chart===match,worksheetDrawingCount:drawings.length,workbook:{available:!!wb,methods:methods(wb),hasGetDrawingsByName:has(wb,'GetDrawingsByName'),byName:byDesc}}
 }catch(e){return {ok:false,source:'live-coedit-editor',error:String(e&&e.message?e.message:e)}}
}
async function run(frame,apiHely,sheet,range='XFA1:XFB4',timeoutMs=15000){const body=`return (${probe.toString()})(${JSON.stringify({sheet,range})});`;return frame.evaluate(({u,body,timeout})=>new Promise(resolve=>{const editor=u==='window.editor'?window.editor:(window.Asc||{}).editor;let done=false;const finish=v=>{if(!done){done=true;resolve(v)}};try{editor.callCommand(new Function(body),false,v=>finish(v))}catch(e){finish({ok:false,error:String(e&&e.message?e.message:e)})}setTimeout(()=>finish({ok:false,error:'callback-timeout'}),timeout)}),{u:apiHely,body,timeout:timeoutMs})}
module.exports={probe,run}
