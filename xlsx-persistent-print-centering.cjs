'use strict'
const persistent=require('./xlsx-persistent-session.cjs')
const LIVE='live-coedit-editor'
function validSpec(s){return s&&typeof s.sheet==='string'&&s.sheet.trim()&&['horizontal','vertical'].some(k=>s[k]!==undefined)&&['horizontal','vertical'].every(k=>s[k]===undefined||typeof s[k]==='boolean')}
function command(spec){
 try{
  var sheets=Api.GetSheets(),a=null;for(var i=0;i<sheets.length;i++)if(sheets[i]&&sheets[i].GetName&&sheets[i].GetName()===spec.sheet){a=sheets[i];break}
  var ws=a&&a.worksheet,po=ws&&ws.PagePrintOptions;if(!a||!ws||!po)return {ok:false,outcome:'print-centering-sheet-or-api-unavailable',source:'live-coedit-editor'}
  function read(){return {horizontal:po.asc_getHorizontalCentered(),vertical:po.asc_getVerticalCentered()}}
  var before=read(),want={horizontal:spec.horizontal===undefined?before.horizontal:spec.horizontal,vertical:spec.vertical===undefined?before.vertical:spec.vertical}
  function eq(x){return x.horizontal===want.horizontal&&x.vertical===want.vertical}
  if(eq(before))return {ok:true,outcome:'print-centering-already-satisfied',source:'live-coedit-editor',noOp:true,applied:false,state:before,verification:{measurable:true,match:true,expected:want}}
  if(!spec.apply)return {ok:true,outcome:'print-centering-observed',source:'live-coedit-editor',noOp:false,applied:false,state:before,verification:{measurable:true,match:false,expected:want}}
  if(spec.horizontal!==undefined)po.asc_setHorizontalCentered(spec.horizontal);if(spec.vertical!==undefined)po.asc_setVerticalCentered(spec.vertical)
  var state=read(),ok=eq(state);return {ok,outcome:ok?'print-centering-applied':'print-centering-verification-mismatch',source:'live-coedit-editor',noOp:false,applied:true,state,verification:{measurable:true,match:ok,expected:want}}
 }catch(err){return {ok:false,outcome:'print-centering-error',source:'live-coedit-editor',error:String(err&&err.message||err)}}
}
async function runCommand(session,spec,apply){if(!validSpec(spec))return {ok:false,outcome:'print-centering-invalid-spec',source:LIVE};const body=`return (${command.toString()})(${JSON.stringify({...spec,apply:!!apply})});`;return session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;e.callCommand(new Function(body),false,resolve)}),{where:session.apiWhere,body})}
function authorityFor(_s,r){return r?.ok&&r?.source===LIVE?'LIVE_VERIFY':r?.source===LIVE?'LIVE_READ':'PLAN_ONLY'}
module.exports={validSpec,command,runCommand,authorityFor}
