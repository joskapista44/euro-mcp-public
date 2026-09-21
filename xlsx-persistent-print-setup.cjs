'use strict'
const persistent=require('./xlsx-persistent-session.cjs')
const LIVE='live-coedit-editor'
function validSpec(s){return s&&typeof s.sheet==='string'&&s.sheet.trim()&&['fit_to_pages','scale','page_size'].includes(s.mode)&&(
 s.mode==='fit_to_pages'?Number.isInteger(s.fitToWidth)&&s.fitToWidth>=0&&Number.isInteger(s.fitToHeight)&&s.fitToHeight>=0:
 s.mode==='scale'?Number.isFinite(s.scale)&&s.scale>=10&&s.scale<=400:
 Number.isFinite(s.width)&&s.width>0&&Number.isFinite(s.height)&&s.height>0)}
function printSetupCommand(spec){
 try{
  var a=null,sheets=Api.GetSheets();for(var si=0;si<sheets.length;si++){var candidate=sheets[si];if(candidate&&candidate.GetName&&candidate.GetName()===spec.sheet){a=candidate;break}}var ws=a&&a.worksheet,po=ws&&ws.PagePrintOptions,ps=po&&po.asc_getPageSetup?po.asc_getPageSetup():null
  if(!a||!ws)return {ok:false,outcome:'print-setup-sheet-unavailable',source:'live-coedit-editor'}
  if(!po||!ps)return {ok:false,outcome:'print-setup-api-unavailable',source:'live-coedit-editor'}
  function read(){return {width:ps.asc_getWidth(),height:ps.asc_getHeight(),fitToWidth:ps.asc_getFitToWidth(),fitToHeight:ps.asc_getFitToHeight(),scale:ps.asc_getScale(),orientation:ps.asc_getOrientation()}}
  var before=read(),want={width:before.width,height:before.height,fitToWidth:before.fitToWidth,fitToHeight:before.fitToHeight,scale:before.scale,orientation:before.orientation}
  if(spec.mode==='fit_to_pages'){want.fitToWidth=spec.fitToWidth;want.fitToHeight=spec.fitToHeight}
  else if(spec.mode==='scale')want.scale=spec.scale
  else {want.width=spec.width;want.height=spec.height}
  function match(x){if(spec.mode==='fit_to_pages')return x.fitToWidth===want.fitToWidth&&x.fitToHeight===want.fitToHeight;if(spec.mode==='scale')return Math.abs(x.scale-want.scale)<.01;return Math.abs(x.width-want.width)<.01&&Math.abs(x.height-want.height)<.01}
  if(match(before))return {ok:true,outcome:'print-setup-already-satisfied',source:'live-coedit-editor',noOp:true,applied:false,before:before,state:before,verification:{measurable:true,match:true,expected:want}}
  if(!spec.apply)return {ok:true,outcome:'print-setup-observed',source:'live-coedit-editor',noOp:false,applied:false,before:before,state:before,verification:{measurable:true,match:false,expected:want}}
  if(spec.mode==='fit_to_pages'){if(typeof ps.asc_setFitToWidth!=='function'||typeof ps.asc_setFitToHeight!=='function')return {ok:false,outcome:'print-setup-fit-api-unavailable',source:'live-coedit-editor'};ps.asc_setFitToWidth(spec.fitToWidth);ps.asc_setFitToHeight(spec.fitToHeight);if(typeof ws.setFitToPage==='function')ws.setFitToPage(spec.fitToWidth,spec.fitToHeight)}
  else if(spec.mode==='scale'){if(typeof ps.asc_setScale!=='function')return {ok:false,outcome:'print-setup-scale-api-unavailable',source:'live-coedit-editor'};ps.asc_setScale(spec.scale)}
  else {if(typeof ps.asc_setWidth!=='function'||typeof ps.asc_setHeight!=='function')return {ok:false,outcome:'print-setup-size-api-unavailable',source:'live-coedit-editor'};ps.asc_setWidth(spec.width);ps.asc_setHeight(spec.height)}
  var state=read(),ok=match(state);return {ok:ok,outcome:ok?'print-setup-applied':'print-setup-verification-mismatch',source:'live-coedit-editor',noOp:false,applied:true,before:before,state:state,verification:{measurable:true,match:ok,expected:want}}
 }catch(err){return {ok:false,outcome:'print-setup-error',source:'live-coedit-editor',error:String(err&&err.message||err)}}
}
async function runCommand(session,spec,apply){if(!validSpec(spec))return {ok:false,outcome:'print-setup-invalid-spec',source:LIVE};const body=`return (${printSetupCommand.toString()})(${JSON.stringify({...spec,apply:!!apply})});`;return session.frame.evaluate(({where,body})=>new Promise(resolve=>{const e=where==='window.editor'?window.editor:(window.Asc||{}).editor;e.callCommand(new Function(body),false,resolve)}),{where:session.apiWhere,body})}
function authorityFor(_s,r){return r?.ok&&r?.source===LIVE?'LIVE_VERIFY':r?.source===LIVE?'LIVE_READ':'PLAN_ONLY'}
async function executePrintSetupTaskInPersistentSession(options={}){const op=options.task?.operations?.length===1?options.task.operations[0]:null;if(!op||op.intent!=='set_print_setup'||!validSpec(op))return {ok:false,outcome:'xlsx-print-setup-task-invalid',authority:'PLAN_ONLY',writeAllowed:false};return persistent.withPersistentXlsxSession(options,async api=>{const pre=await runCommand(api.session,op,false);if(pre?.ok&&pre.noOp)return {ok:true,outcome:'xlsx-print-setup-task-already-satisfied',authority:'LIVE_VERIFY',noOp:true,wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:pre.state}};const applied=await runCommand(api.session,op,true);if(!applied?.ok||!applied.verification?.match)return {ok:false,outcome:'xlsx-print-setup-task-live-verify-failed',authority:'LIVE_READ',writeAllowed:false,applied};api.session.markWrite();return {ok:true,outcome:'xlsx-print-setup-task-live-verified',authority:'LIVE_VERIFY',noOp:false,wholeTaskVerification:{ok:true,authority:'LIVE_VERIFY',state:applied.state}}})}
module.exports={validSpec,printSetupCommand,runCommand,authorityFor,executePrintSetupTaskInPersistentSession}
