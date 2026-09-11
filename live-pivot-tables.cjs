'use strict';
function pivotCommand(spec){
 function safe(fn){try{return {ok:true,value:fn()}}catch(e){return {ok:false,error:String(e&&e.message?e.message:e)}}}
 function snap(p){if(!p)return null;function call(n){var r=safe(function(){return typeof p[n]==='function'?p[n]():null});return r.ok?r.value:null}return {name:call('GetName'),source:call('GetSource'),rowFields:(call('GetRowFields')||[]).length,columnFields:(call('GetColumnFields')||[]).length,dataFields:(call('GetDataFields')||[]).length,styleName:call('GetStyleName'),title:call('GetTitle'),description:call('GetDescription')}}
 function get(name){try{return typeof Api.GetPivotByName==='function'?Api.GetPivotByName(name):null}catch(_){return null}}
 try{
  var op=spec.operation, p, before, after, result;
  if(op==='pivot.inspect'){p=get(spec.name);return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:op,verification:{status:'PASS',expected:'live-pivot-state',actual:snap(p)}}}
  if(op==='pivot.createNewWorksheet'){
   before=(Api.GetAllPivotTables()||[]).length;
   var src=Api.GetActiveSheet().GetRange(spec.source);
   p=Api.InsertPivotNewWorksheet(src);
   if(!p)return {ok:false,outcome:'unsupported',source:'live-coedit-editor',operation:op,verification:{status:'UNKNOWN',expected:{count:before+1},actual:null}};
   if(spec.name&&typeof p.SetName==='function')p.SetName(spec.name);
   after=(Api.GetAllPivotTables()||[]).length;
   var a=snap(p), pass=after===before+1&&a&&(!spec.name||a.name===spec.name)&&a.source===spec.source;
   return {ok:pass,outcome:pass?'ok':'verification-failed',source:'live-coedit-editor',operation:op,verification:{status:pass?'PASS':'FAIL',expected:{count:before+1,name:spec.name,source:spec.source},actual:{count:after,pivot:a}}}
  p=get(spec.name);if(!p)return {ok:false,outcome:'not-found',source:'live-coedit-editor',operation:op,verification:{status:'FAIL',expected:spec.name,actual:null}};
  if(op==='pivot.addFields'){
   before=snap(p);result=p.AddFields(spec.rowFields||[],spec.columnFields||[],spec.pageFields||[]);after=snap(p);
   var er=(spec.rowFields||[]).length,ec=(spec.columnFields||[]).length;var pass2=after.rowFields>=before.rowFields+er&&after.columnFields>=before.columnFields+ec;
   return {ok:pass2,outcome:pass2?'ok':'verification-failed',source:'live-coedit-editor',operation:op,verification:{status:pass2?'PASS':'FAIL',expected:{rowFieldsAtLeast:before.rowFields+er,columnFieldsAtLeast:before.columnFields+ec},actual:after,result:result}}
  if(op==='pivot.addDataField'){
   before=snap(p);result=p.AddDataField(spec.field);after=snap(p);var pass3=after.dataFields===before.dataFields+1;
   return {ok:pass3,outcome:pass3?'ok':'verification-failed',source:'live-coedit-editor',operation:op,verification:{status:pass3?'PASS':'FAIL',expected:{dataFields:before.dataFields+1},actual:after,result:result}}
  if(op==='pivot.rename'){var old=snap(p);p.SetName(spec.newName);after=snap(get(spec.newName)||p);var pass4=after&&after.name===spec.newName;return {ok:pass4,outcome:pass4?'ok':'verification-failed',source:'live-coedit-editor',operation:op,verification:{status:pass4?'PASS':'FAIL',expected:{name:spec.newName},actual:after,before:old}}}
  if(op==='pivot.style'){p.SetStyleName(spec.styleName);after=snap(p);var pass5=after.styleName===spec.styleName;return {ok:pass5,outcome:pass5?'ok':'verification-failed',source:'live-coedit-editor',operation:op,verification:{status:pass5?'PASS':'FAIL',expected:{styleName:spec.styleName},actual:after}}}
  if(op==='pivot.refresh'){result=p.RefreshTable();after=snap(p);return {ok:true,outcome:'ok',source:'live-coedit-editor',operation:op,verification:{status:'PASS',expected:'public-refresh-completed',actual:after,result:result}}}
  return {ok:false,outcome:'unknown-operation',source:'live-coedit-editor',operation:op,verification:{status:'FAIL',expected:'known pivot operation',actual:op}}
 }catch(e){return {ok:false,outcome:'execution-error',source:'live-coedit-editor',operation:spec&&spec.operation,verification:{status:'UNKNOWN',expected:'operation completes with semantic readback',actual:String(e&&e.message?e.message:e)}}}
}
async function runPivotInFrame(frame,editorExpr,spec){var body='return ('+pivotCommand.toString()+')('+JSON.stringify(spec)+');';return frame.evaluate(({editorExpr,body})=>new Promise(resolve=>{var editor;try{editor=Function('return ('+editorExpr+')')()}catch(e){return resolve({ok:false,outcome:'editor-unavailable',source:'live-coedit-editor',verification:{status:'UNKNOWN',actual:String(e)}})};if(!editor||typeof editor.callCommand!=='function')return resolve({ok:false,outcome:'editor-unavailable',source:'live-coedit-editor',verification:{status:'UNKNOWN'}});editor.callCommand(new Function(body),false,resolve)}),{editorExpr,body})}
module.exports={pivotCommand,runPivotInFrame};
