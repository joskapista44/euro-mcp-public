'use strict'
const LIVE='live-coedit-editor'
function validSpec(s){return s&&typeof s.sheet==='string'&&s.sheet.trim()&&['rows','columns'].includes(s.axis)&&Number.isInteger(s.from)&&s.from>=1&&Number.isInteger(s.to)&&s.to>=s.from}
function command(spec){
 function col(n){var s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26)}return s}
 try{
  var sheets=Api.GetSheets(),a=null,idx=-1;for(var i=0;i<sheets.length;i++)if(sheets[i]&&sheets[i].GetName&&sheets[i].GetName()===spec.sheet){a=sheets[i];idx=i;break}
  if(!a||idx<0)return {ok:false,outcome:'print-titles-sheet-unavailable',source:'live-coedit-editor'}
  var want=spec.axis==='rows'?'$'+spec.from+':$'+spec.to:'$'+col(spec.from)+':$'+col(spec.to)
  var editor=window.Asc&&window.Asc.editor;if(!editor||typeof editor.asc_getPageOptions!=='function'||typeof editor.asc_changePrintTitles!=='function')return {ok:false,outcome:'print-titles-editor-api-unavailable',source:'live-coedit-editor'}
  function read(){var po=editor.asc_getPageOptions(idx,true,false);return {rows:po&&po.asc_getPrintTitlesHeight?po.asc_getPrintTitlesHeight():po&&po.printTitlesHeight||null,columns:po&&po.asc_getPrintTitlesWidth?po.asc_getPrintTitlesWidth():po&&po.printTitlesWidth||null}}
  var before=read(),actual=spec.axis==='rows'?before.rows:before.columns
  if(actual===want)return {ok:true,outcome:'print-titles-already-satisfied',source:'live-coedit-editor',noOp:true,applied:false,state:before,verification:{measurable:true,match:true,expected:want}}
  if(!spec.apply)return {ok:true,outcome:'print-titles-observed',source:'live-coedit-editor',noOp:false,applied:false,state:before,verification:{measurable:true,match:false,expected:want,actual:actual}}
  var cols=spec.axis==='columns'?want:before.columns,rows=spec.axis==='rows'?want:before.rows
  var ret=editor.asc_changePrintTitles(cols,rows,idx)
  if(ret===false)return {ok:false,outcome:'print-titles-editor-mutation-rejected',source:'live-coedit-editor'}
  var state=read(),now=spec.axis==='rows'?state.rows:state.columns,ok=now===want
  return {ok:ok,outcome:ok?'print-titles-applied':'print-titles-semantic-mismatch',source:'live-coedit-editor',noOp:false,applied:true,state:state,verification:{measurable:true,match:ok,expected:want,actual:now}}
 }catch(err){return {ok:false,outcome:'print-titles-error',source:'live-coedit-editor',error:String(err&&err.message||err)}}
}
async function immediate(session,spec,apply){if(!validSpec(spec))return {ok:false,outcome:'print-titles-invalid-spec',source:LIVE};const body=`return (${command.toString()})(${JSON.stringify({...spec,apply:!!apply})});`;return session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;e.callCommand(new Function(body),false,resolve)}),{where:session.apiWhere,body})}
async function runCommand(session,spec,apply,{timeoutMs=15000,pollMs=25}={}){let r=await immediate(session,spec,apply);if(!(apply&&r?.applied===true&&!r?.ok))return r;const started=Date.now();let reads=0,last=r;while(Date.now()-started<=timeoutMs){last=await immediate(session,spec,false);reads++;if(last?.ok&&last?.noOp)return {...r,ok:true,outcome:'print-titles-live-verified',state:last.state,verification:{measurable:true,match:true},collaborationReadback:{reads,waitMs:Date.now()-started}};if(Date.now()-started>=timeoutMs)break;await new Promise(x=>setTimeout(x,pollMs))}return {...r,collaborationReadback:{reads,waitMs:Date.now()-started,state:last?.state||null}}}
function authorityFor(_s,r){return r?.ok&&r?.source===LIVE?'LIVE_VERIFY':r?.source===LIVE?'LIVE_READ':'PLAN_ONLY'}
module.exports={validSpec,runCommand,authorityFor}
