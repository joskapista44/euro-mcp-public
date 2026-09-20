'use strict'
const LIVE='live-coedit-editor'
function validSpec(s){return s&&typeof s.sheet==='string'&&s.sheet.trim()&&['rows','columns'].includes(s.axis)&&Number.isInteger(s.from)&&s.from>=1&&Number.isInteger(s.to)&&s.to>=s.from}
function command(spec){
 function col(n){var s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
 function escSheet(n){return String(n).replace(/'/g,"''")}
 try{
  var sheets=Api.GetSheets(),a=null;for(var i=0;i<sheets.length;i++)if(sheets[i]&&sheets[i].GetName&&sheets[i].GetName()===spec.sheet){a=sheets[i];break}
  var ws=a&&a.worksheet,po=ws&&ws.PagePrintOptions;if(!a||!ws||!po)return {ok:false,outcome:'print-titles-sheet-or-api-unavailable',source:'live-coedit-editor'}
  var want=spec.axis==='rows'?'$'+spec.from+':$'+spec.to:'$'+col(spec.from)+':$'+col(spec.to)
  function read(){if(typeof po.initPrintTitles==='function')po.initPrintTitles();var d=null;try{d=Api.GetDefName('Print_Titles')}catch(_){}
   return {rows:po.asc_getPrintTitlesHeight(),columns:po.asc_getPrintTitlesWidth(),definedName:d&&d.GetRefersTo?d.GetRefersTo():null}}
  function desiredRef(state){var rows=spec.axis==='rows'?want:state.rows,cols=spec.axis==='columns'?want:state.columns,parts=[];if(cols)parts.push("='"+escSheet(spec.sheet)+"'!"+cols);if(rows)parts.push("='"+escSheet(spec.sheet)+"'!"+rows);return parts.join(',')}
  var before=read(),actual=spec.axis==='rows'?before.rows:before.columns,wantRef=desiredRef(before)
  if(actual===want&&before.definedName===wantRef)return {ok:true,outcome:'print-titles-already-satisfied',source:'live-coedit-editor',noOp:true,applied:false,state:before,verification:{measurable:true,match:true,expected:want,expectedRef:wantRef}}
  if(!spec.apply)return {ok:true,outcome:'print-titles-observed',source:'live-coedit-editor',noOp:false,applied:false,state:before,verification:{measurable:true,match:false,expected:want,expectedRef:wantRef}}
  var d=null;try{d=Api.GetDefName('Print_Titles')}catch(_){}
  if(d&&typeof d.SetRefersTo==='function')d.SetRefersTo(wantRef);else if(typeof Api.AddDefName==='function')Api.AddDefName('Print_Titles',wantRef);else return {ok:false,outcome:'print-titles-defined-name-api-unavailable',source:'live-coedit-editor'}
  if(spec.axis==='rows')po.asc_setPrintTitlesHeight(want);else po.asc_setPrintTitlesWidth(want)
  var state={rows:po.asc_getPrintTitlesHeight(),columns:po.asc_getPrintTitlesWidth(),definedName:null};try{var dn=Api.GetDefName('Print_Titles');state.definedName=dn&&dn.GetRefersTo?dn.GetRefersTo():null}catch(_){}
  var now=spec.axis==='rows'?state.rows:state.columns,ok=now===want&&state.definedName===wantRef
  return {ok:ok,outcome:ok?'print-titles-applied':'print-titles-semantic-mismatch',source:'live-coedit-editor',noOp:false,applied:true,state:state,verification:{measurable:true,match:ok,expected:want,expectedRef:wantRef}}
 }catch(err){return {ok:false,outcome:'print-titles-error',source:'live-coedit-editor',error:String(err&&err.message||err)}}
}
async function immediate(session,spec,apply){if(!validSpec(spec))return {ok:false,outcome:'print-titles-invalid-spec',source:LIVE};const body=`return (${command.toString()})(${JSON.stringify({...spec,apply:!!apply})});`;return session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;e.callCommand(new Function(body),false,resolve)}),{where:session.apiWhere,body})}
async function runCommand(session,spec,apply,{timeoutMs=15000,pollMs=25}={}){let r=await immediate(session,spec,apply);if(!(apply&&r?.applied===true&&!r?.ok))return r;const started=Date.now();let reads=0,last=r;while(Date.now()-started<=timeoutMs){last=await immediate(session,spec,false);reads++;if(last?.ok&&last?.noOp)return {...r,ok:true,outcome:'print-titles-live-verified',state:last.state,verification:{measurable:true,match:true},collaborationReadback:{reads,waitMs:Date.now()-started}};if(Date.now()-started>=timeoutMs)break;await new Promise(x=>setTimeout(x,pollMs))}return {...r,collaborationReadback:{reads,waitMs:Date.now()-started,state:last?.state||null}}}
function authorityFor(_s,r){return r?.ok&&r?.source===LIVE?'LIVE_VERIFY':r?.source===LIVE?'LIVE_READ':'PLAN_ONLY'}
module.exports={validSpec,runCommand,authorityFor}
