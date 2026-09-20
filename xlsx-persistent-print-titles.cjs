'use strict'
const persistent=require('./xlsx-persistent-session.cjs')
const LIVE='live-coedit-editor'
function validSpec(s){return s&&typeof s.sheet==='string'&&s.sheet.trim()&&['rows','columns'].includes(s.axis)&&Number.isInteger(s.from)&&s.from>=1&&Number.isInteger(s.to)&&s.to>=s.from}
function absTitle(axis,from,to){return axis==='rows'?`$1:$${to}`.replace('$1',`$${from}`):`$A:$${String.fromCharCode(64+to)}`}
function command(spec){
 try{
  var sheets=Api.GetSheets(),a=null;for(var i=0;i<sheets.length;i++)if(sheets[i]&&sheets[i].GetName&&sheets[i].GetName()===spec.sheet){a=sheets[i];break}
  var ws=a&&a.worksheet,po=ws&&ws.PagePrintOptions;if(!a||!ws||!po)return {ok:false,outcome:'print-titles-sheet-or-api-unavailable',source:'live-coedit-editor'}
  function col(n){var s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
  var want=spec.axis==='rows'?'$'+spec.from+':$'+spec.to:'$'+col(spec.from)+':$'+col(spec.to)
  function read(){return {rows:po.asc_getPrintTitlesHeight(),columns:po.asc_getPrintTitlesWidth()}}
  var before=read(),actual=spec.axis==='rows'?before.rows:before.columns
  if(actual===want)return {ok:true,outcome:'print-titles-already-satisfied',source:'live-coedit-editor',noOp:true,applied:false,state:before,verification:{measurable:true,match:true,expected:want}}
  if(!spec.apply)return {ok:true,outcome:'print-titles-observed',source:'live-coedit-editor',noOp:false,applied:false,state:before,verification:{measurable:true,match:false,expected:want}}
  if(spec.axis==='rows')po.asc_setPrintTitlesHeight(want);else po.asc_setPrintTitlesWidth(want)
  var state=read(),now=spec.axis==='rows'?state.rows:state.columns,ok=now===want
  return {ok:ok,outcome:ok?'print-titles-applied':'print-titles-verification-mismatch',source:'live-coedit-editor',noOp:false,applied:true,state:state,verification:{measurable:true,match:ok,expected:want,actual:now}}
 }catch(err){return {ok:false,outcome:'print-titles-error',source:'live-coedit-editor',error:String(err&&err.message||err)}}
}
async function runCommand(session,spec,apply){if(!validSpec(spec))return {ok:false,outcome:'print-titles-invalid-spec',source:LIVE};const body=`return (${command.toString()})(${JSON.stringify({...spec,apply:!!apply})});`;return session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;e.callCommand(new Function(body),false,resolve)}),{where:session.apiWhere,body})}
function authorityFor(_s,r){return r?.ok&&r?.source===LIVE?'LIVE_VERIFY':r?.source===LIVE?'LIVE_READ':'PLAN_ONLY'}
module.exports={validSpec,runCommand,authorityFor}
